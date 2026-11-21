"""KRONOS Offline version 6, part 6b (D119): a changed copy that keeps everything.

When a recipe moved no row and no sheet, the changed copy is the uploaded file itself with only the changed cells
written into its XML: values, formulas, fills and number formats (a new cell style beside the cell's own, so every
other cell keeps its look), and hidden rows. Every other part of the file stays byte for byte, so pictures, charts,
print areas, page setup, merged cells, data validation, notes and macros all survive. Excel recalculates every
formula on opening (fullCalcOnLoad), and when a formula cell changed the calculation chain is dropped, which Excel
builds again.

    patch(loaded, model) -> the copy's bytes, or None when the change needs the full copy (kronos_book's openpyxl one)

None when: a sheet was added, removed or reordered; a row moved, was inserted or deleted (a cell's origin isn't where
it is); a changed cell is the first cell of a shared formula or part of an array formula; a cell whose style a CLEAR
dropped holds a note or a link (the full copy drops those with it).
"""
import copy
import datetime as dt
import io
import re
import zipfile
import xml.etree.ElementTree as ET

import semantics as S
from vcl_docx import xml_parse, xml_write
from vcl_fixtures import Formula, col_letter, col_number
from vcl_template import BUILTIN, DATE_IDS, is_date_format

M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
EPOCH = dt.datetime(1899, 12, 30)
REF = re.compile(r"([A-Z]+)(\d+)")


class _Cant(Exception):
    """This change needs the full copy."""


def _rc(ref):
    m = REF.fullmatch(ref)
    return int(m.group(2)), col_number(m.group(1))


def _sheet_parts(parts):
    wb = ET.fromstring(parts["xl/workbook.xml"])
    rels = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    target = {r.get("Id"): r.get("Target") for r in rels.iter(PKG + "Relationship")}
    out = {}
    for s in wb.iter(M + "sheet"):
        t = target.get(s.get(R + "id"), "").lstrip("/")
        out[s.get("name")] = t if t.startswith("xl/") else "xl/" + t
    return out


def _same(a, b):
    if isinstance(a, bool) != isinstance(b, bool):
        return False
    return a == b


class _Styles:
    """xl/styles.xml, with new cell styles added as needed (each a copy of the cell's own with another fill or number
    format)."""

    def __init__(self, data):
        self.original = data
        self.root = xml_parse(data)
        self.xfs = self.root.find(M + "cellXfs")
        if self.xfs is None:
            raise _Cant("no cell styles")
        self.fills = self.root.find(M + "fills")
        self.numfmts = self.root.find(M + "numFmts")
        self.cache, self.changed = {}, False

    def code_of(self, s):
        xf = list(self.xfs)[s] if s < len(self.xfs) else list(self.xfs)[0]
        num = int(xf.get("numFmtId", "0"))
        if self.numfmts is not None:
            for f in self.numfmts:
                if int(f.get("numFmtId")) == num:
                    return num, f.get("formatCode")
        return num, BUILTIN.get(num, "General")

    def is_date(self, s):
        num, code = self.code_of(s)
        return num in DATE_IDS or (num >= 164 and is_date_format(code))

    def _numfmt_id(self, code):
        for num, text in BUILTIN.items():
            if text == code:
                return num
        if self.numfmts is None:
            self.numfmts = ET.Element(M + "numFmts")
            self.root.insert(0, self.numfmts)
        ids = [int(f.get("numFmtId")) for f in self.numfmts]
        for f in self.numfmts:
            if f.get("formatCode") == code:
                return int(f.get("numFmtId"))
        new = max([163] + ids) + 1
        ET.SubElement(self.numfmts, M + "numFmt", {"numFmtId": str(new), "formatCode": code})
        self.numfmts.set("count", str(len(self.numfmts)))
        return new

    def _fill_id(self, fill):
        if fill is None:
            return 0
        if self.fills is None:
            raise _Cant("no fills")
        rgb = "FF" + fill.lstrip("#").upper()
        for i, f in enumerate(self.fills):
            p = f.find(M + "patternFill")
            fg = p.find(M + "fgColor") if p is not None else None
            if p is not None and p.get("patternType") == "solid" and fg is not None and fg.get("rgb", "").upper() == rgb:
                return i
        new = ET.SubElement(self.fills, M + "fill")
        p = ET.SubElement(new, M + "patternFill", {"patternType": "solid"})
        ET.SubElement(p, M + "fgColor", {"rgb": rgb})
        ET.SubElement(p, M + "bgColor", {"indexed": "64"})
        self.fills.set("count", str(len(self.fills)))
        return len(self.fills) - 1

    def style(self, base, fill, code):
        """The index of a cell style like `base` with this fill (None: none) and number format (None: base's own)."""
        key = (base, fill, code)
        if key not in self.cache:
            xf = copy.deepcopy(list(self.xfs)[base] if base < len(self.xfs) else list(self.xfs)[0])
            fill_id = self._fill_id(fill)
            xf.set("fillId", str(fill_id))
            xf.set("applyFill", "1")
            if code is not None:
                xf.set("numFmtId", str(self._numfmt_id(code)))
                xf.set("applyNumberFormat", "1")
            self.xfs.append(xf)
            self.xfs.set("count", str(len(self.xfs)))
            self.cache[key] = len(self.xfs) - 1
            self.changed = True
        return self.cache[key]


def _write_value(c, value, r, col, styles):
    """Put the interpreter's value into the <c> element: a number, text (inline), TRUE/FALSE, an error, a date (as its
    serial number), a formula, or nothing."""
    for child in list(c):
        if child.tag in (M + "f", M + "v", M + "is"):
            c.remove(child)
    c.attrib.pop("t", None)
    if value is None:
        return
    if isinstance(value, Formula):
        from kronos_book import r1c1_to_a1
        text = value.text if not value.r1c1 else r1c1_to_a1(value.text, r, col)
        ET.SubElement(c, M + "f").text = text[1:] if text.startswith("=") else text
        return
    if isinstance(value, bool):
        c.set("t", "b")
        ET.SubElement(c, M + "v").text = "1" if value else "0"
    elif isinstance(value, S.XlError):
        c.set("t", "e")
        ET.SubElement(c, M + "v").text = value.text
    elif isinstance(value, (int, float)):
        ET.SubElement(c, M + "v").text = repr(float(value)) if isinstance(value, float) else str(value)
    elif isinstance(value, (dt.date, dt.datetime)):
        when = value if isinstance(value, dt.datetime) else dt.datetime(value.year, value.month, value.day)
        serial = (when - EPOCH).total_seconds() / 86400
        ET.SubElement(c, M + "v").text = str(int(serial)) if serial == int(serial) else repr(serial)
    else:
        c.set("t", "inlineStr")
        t = ET.SubElement(ET.SubElement(c, M + "is"), M + "t")
        t.text = str(value)
        t.set(XML_SPACE, "preserve")


def _cell(rows, sheet_data, r, col):
    """The <c> for (r, col), made (in column order, in its row, made in row order) when it isn't there."""
    row = rows.get(r)
    if row is None:
        row = ET.Element(M + "row", {"r": str(r)})
        later = [x for x in sheet_data if x.tag == M + "row" and int(x.get("r")) > r]
        sheet_data.insert(list(sheet_data).index(later[0]) if later else len(sheet_data), row)
        rows[r] = row
    ref = f"{col_letter(col)}{r}"
    for c in row.findall(M + "c"):
        if c.get("r") == ref:
            return c
    c = ET.Element(M + "c", {"r": ref})
    later = [x for x in row.findall(M + "c") if _rc(x.get("r"))[1] > col]
    row.insert(list(row).index(later[0]) if later else len(row), c)
    row.attrib.pop("spans", None)   # the row's column span is only a hint; Excel works it out again
    return c


def _row(rows, sheet_data, r):
    if r not in rows:
        row = ET.Element(M + "row", {"r": str(r)})
        later = [x for x in sheet_data if x.tag == M + "row" and int(x.get("r")) > r]
        sheet_data.insert(list(sheet_data).index(later[0]) if later else len(sheet_data), row)
        rows[r] = row
    return rows[r]


def _check_unmoved(loaded, model):
    if list(model.sheets) != list(loaded.initial):
        raise _Cant("a sheet was added, removed or reordered")
    for name, sh in model.sheets.items():
        initial = loaded.initial[name]
        before, after = initial.origins or {}, sh.origins or {}
        if set(before) != set(after):
            raise _Cant("a row was inserted or deleted")
        for (r, col), o in after.items():
            if o[0] != name or (o[1], o[2]) != (r, col):
                raise _Cant("a row moved")


def patch(loaded, model):
    """The uploaded file with only the changed cells rewritten, or None when this change needs the full copy."""
    try:
        return _patch(loaded, model)
    except (_Cant, KeyError, ValueError, ET.ParseError):
        return None


def _patch(loaded, model):
    _check_unmoved(loaded, model)
    with zipfile.ZipFile(io.BytesIO(loaded.data)) as z:
        infos = z.infolist()
        parts = {i.filename: z.read(i.filename) for i in infos}
    where = _sheet_parts(parts)
    styles = _Styles(parts["xl/styles.xml"]) if "xl/styles.xml" in parts else None
    out = {}
    values_changed = formulas_changed = False
    for name, sh in model.sheets.items():
        initial = loaded.initial[name]
        changed_cells = [rc for rc in set(sh.cells) | set(initial.cells)
                         if not (_same(initial.cells.get(rc), sh.cells.get(rc))
                                 or (isinstance(initial.cells.get(rc), Formula) and initial.cells.get(rc) == sh.cells.get(rc)))]
        restyled = [rc for rc in set(sh.fills) | set(initial.fills) | set(sh.formats) | set(initial.formats)
                    if sh.fills.get(rc) != initial.fills.get(rc) or sh.formats.get(rc) != initial.formats.get(rc)]
        dropped = [rc for rc, o in (sh.origins or {}).items() if not o[3]]
        hidden = sh.hidden ^ initial.hidden
        if not (changed_cells or restyled or dropped or hidden):
            continue
        part = where[name]
        root = xml_parse(parts[part])
        sheet_data = root.find(M + "sheetData")
        if any(x.get("r") is None for x in sheet_data.findall(M + "row")) or \
                any(c.get("r") is None for x in sheet_data.findall(M + "row") for c in x.findall(M + "c")):
            raise _Cant("rows or cells without their own reference")
        rows = {int(x.get("r")): x for x in sheet_data.findall(M + "row")}
        kept = loaded.cells.get(name, {})
        for rc in dropped:   # CLEAR formats or all: the cell loses its own style, as the full copy does
            extras = kept.get(rc, (None, None, None, None, (None, None)))[4]
            if extras[0] is not None or extras[1] is not None:
                raise _Cant("a cleared cell holds a note or a link")
        for rc in sorted(set(changed_cells) | set(restyled) | set(dropped)):
            r, col = rc
            c = _cell(rows, sheet_data, r, col)
            f = c.find(M + "f")
            if rc in changed_cells and f is not None and (f.get("t") == "array" or (f.get("t") == "shared" and f.get("ref"))):
                raise _Cant("the first cell of a shared formula, or an array formula")
            if rc in changed_cells:
                values_changed = True
                if f is not None or isinstance(sh.cells.get(rc), Formula):
                    formulas_changed = True
                _write_value(c, sh.cells.get(rc), r, col, styles)
            value = sh.cells.get(rc)
            base = 0 if rc in dropped else int(c.get("s", "0"))
            want_fill = sh.fills.get(rc)
            want_fmt = sh.formats.get(rc)
            had_fill = None if rc in dropped else initial.fills.get(rc)
            had_fmt = None if rc in dropped else initial.formats.get(rc)
            needs_date = isinstance(value, (dt.date, dt.datetime)) and styles is not None and not styles.is_date(base) \
                and want_fmt is None
            if styles is not None and (want_fill != had_fill or want_fmt != had_fmt or needs_date or rc in dropped):
                code = want_fmt if want_fmt is not None else ("yyyy-mm-dd" if needs_date and not isinstance(value, dt.datetime)
                                                             else "yyyy-mm-dd hh:mm:ss" if needs_date else None)
                if code is None and rc in dropped:
                    code = "General"
                if want_fill == had_fill and code is None and rc not in dropped:
                    continue
                s = styles.style(base, want_fill, code)
                c.set("s", str(s))
        for r in sorted(hidden):
            row = _row(rows, sheet_data, r)
            if r in sh.hidden:
                row.set("hidden", "1")
            else:
                row.attrib.pop("hidden", None)
        _dimension(root, sheet_data)
        out[part] = xml_write(root, parts[part])
    if not out and not (styles and styles.changed):
        return loaded.data   # nothing to write: the very same file
    if styles is not None and styles.changed:
        out["xl/styles.xml"] = xml_write(styles.root, styles.original)
    drop = set()
    if values_changed:
        wb = xml_parse(parts["xl/workbook.xml"])
        calc = wb.find(M + "calcPr")
        if calc is None:
            calc = ET.Element(M + "calcPr", {"calcId": "0"})
            after = [x for x in wb if x.tag in (M + "definedNames", M + "externalReferences", M + "functionGroups",
                                                 M + "sheets")]
            wb.insert(list(wb).index(after[-1]) + 1 if after else len(wb), calc)
        calc.set("fullCalcOnLoad", "1")
        out["xl/workbook.xml"] = xml_write(wb, parts["xl/workbook.xml"])
    if formulas_changed and "xl/calcChain.xml" in parts:   # Excel builds the chain again
        drop.add("xl/calcChain.xml")
        rels = xml_parse(parts["xl/_rels/workbook.xml.rels"])
        for rel in list(rels):
            if rel.get("Target", "").lstrip("/").replace("xl/", "") == "calcChain.xml":
                rels.remove(rel)
        out["xl/_rels/workbook.xml.rels"] = xml_write(rels, parts["xl/_rels/workbook.xml.rels"])
        types = xml_parse(parts["[Content_Types].xml"])
        for o in list(types):
            if o.get("PartName") == "/xl/calcChain.xml":
                types.remove(o)
        out["[Content_Types].xml"] = xml_write(types, parts["[Content_Types].xml"])
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for info in infos:
            if info.filename in drop:
                continue
            new = zipfile.ZipInfo(info.filename, info.date_time)
            new.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(new, out.get(info.filename, parts[info.filename]))
    return buf.getvalue()


def _dimension(root, sheet_data):
    """The sheet's used area, as the cells now say."""
    dim = root.find(M + "dimension")
    refs = [_rc(c.get("r")) for row in sheet_data.findall(M + "row") for c in row.findall(M + "c") if c.get("r")]
    if dim is None or not refs:
        return
    r1, r2 = min(r for r, _ in refs), max(r for r, _ in refs)
    c1, c2 = min(c for _, c in refs), max(c for _, c in refs)
    dim.set("ref", f"{col_letter(c1)}{r1}" + ("" if (r1, c1) == (r2, c2) else f":{col_letter(c2)}{r2}"))
