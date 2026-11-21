"""Template FIXTUREs (VCL 0.7 part 3a, docs/vep-0.7-documents.md, D110): a practice workbook read from a look-alike
template in data/templates/, with the standard library only (zipfile and xml.etree), so the toolchain and KRONOS
Offline's page need nothing new.

    template_path(name)  -> the template's path (data/templates/<name>.xlsx)
    load(data)           -> {sheet name: Sheet}, in the workbook's order

What a sheet holds: each cell's value (text, number, TRUE/FALSE, a date for a number with a date format); each
formula as a Formula with its text ("=" first) and its saved value; number formats other than General; solid fills;
merged ranges (Sheet.merged, as (row1, col1, row2, col2)). Fonts, borders, widths, pictures, print areas and comments
are not part of the model. Refused, with a TemplateError: shared and array formulas, a link to another workbook, a cell
holding an error value, and a file that isn't an .xlsx workbook.
"""
import datetime as dt
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from errcodes import coded
from resources import ROOT
from semantics import XlError
from vcl_fixtures import Formula, Sheet, col_number

TEMPLATES = ROOT / "data" / "templates"   # KRONOS.exe and KRONOS Offline ship this folder too (D118)
NAME = re.compile(r"[A-Za-z0-9_]{1,60}")
M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG = "{http://schemas.openxmlformats.org/package/2006/relationships}"
EPOCH = dt.datetime(1899, 12, 30)
DATE_IDS = set(range(14, 23)) | set(range(27, 37)) | set(range(45, 48)) | set(range(50, 59))
BUILTIN = {0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%", 11: "0.00E+00",
           12: "# ?/?", 13: "# ??/??", 14: "mm-dd-yy", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM",
           19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss", 22: "m/d/yy h:mm", 37: "#,##0 ;(#,##0)",
           38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)", 45: "mm:ss",
           46: "[h]:mm:ss", 47: "mmss.0", 48: "##0.0E+0", 49: "@"}
CELL_REF = re.compile(r"([A-Z]{1,3})(\d+)")


class TemplateError(Exception):
    """A template the reader can't turn into a practice workbook, in words for the spec's writer."""


def template_path(name):
    """data/templates/<name>.xlsx, for a name of letters, digits and _."""
    if not NAME.fullmatch(name or ""):
        raise TemplateError(coded(f"template={name!r}: a template name is letters, digits and _", "X021"))
    return TEMPLATES / f"{name}.xlsx"


def exists(name):
    try:
        return template_path(name).is_file()
    except TemplateError:
        return False


def is_date_format(code):
    """A number format that shows a date or time: d, m, y, h or s outside quotes, brackets and escapes."""
    text = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]|_.|\*.', "", code or "")
    return bool(re.search(r"[dmyhs]", text, re.I)) and not re.fullmatch(r"(?i)general", text.strip())


def _cell(ref):
    m = CELL_REF.fullmatch(ref.replace("$", ""))
    return int(m.group(2)), col_number(m.group(1))


def _text(e):
    """The text of a shared or inline string: plain <t>, or the runs of rich text."""
    return "".join(t.text or "" for t in e.iter(M + "t"))


def _serial_to_model(serial):
    when = EPOCH + dt.timedelta(days=serial)
    return when.date() if when.time() == dt.time() else when


def _styles(z):
    """[(number format, fill or None)] per cell style index."""
    if "xl/styles.xml" not in z.namelist():
        return []
    root = ET.fromstring(z.read("xl/styles.xml"))
    custom = {int(f.get("numFmtId")): f.get("formatCode") for f in root.iter(M + "numFmt")}
    fills = []
    fills_node = root.find(M + "fills")
    for fill in (fills_node if fills_node is not None else []):
        pattern = fill.find(M + "patternFill")
        rgb = None
        if pattern is not None and pattern.get("patternType") == "solid":
            fg = pattern.find(M + "fgColor")
            if fg is not None and fg.get("rgb") and len(fg.get("rgb")) == 8 and fg.get("rgb").upper() != "00000000":
                rgb = "#" + fg.get("rgb")[2:].upper()
        fills.append(rgb)
    out = []
    xfs = root.find(M + "cellXfs")
    for xf in (xfs if xfs is not None else []):
        num = int(xf.get("numFmtId", "0"))
        code = custom.get(num, BUILTIN.get(num, "General"))
        fill_id = int(xf.get("fillId", "0"))
        out.append((code, fills[fill_id] if fill_id < len(fills) else None, num))
    return out


def _sheet_parts(z):
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    target = {r.get("Id"): r.get("Target") for r in rels.iter(PKG + "Relationship")}
    out = []
    for s in wb.iter(M + "sheet"):
        t = target[s.get(R + "id")].lstrip("/")
        out.append((s.get("name"), t if t.startswith("xl/") else "xl/" + t))
    return out


def load(data, name="template", reading=False):
    """An .xlsx template's bytes -> {sheet name: Sheet}, in the workbook's order.
    reading=True (D118): a workbook a recipe only reads (a TEST's read= one, a DOCUMENT's), as Excel last saved it.
    Every formula then keeps its saved value whatever kind it is (shared, array), an error value is read as one, and a
    link to another workbook is no reason to refuse: only the saved values are read."""
    if not zipfile.is_zipfile(io.BytesIO(data)):
        raise TemplateError(coded(f"{name}: not an .xlsx workbook", "X022"))
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        names = z.namelist()
        if "xl/workbook.xml" not in names:
            raise TemplateError(coded(f"{name}: not an .xlsx workbook", "X023"))
        if not reading and any(n.startswith("xl/externalLinks/") for n in names):
            raise TemplateError(coded(f"{name}: it links to another workbook, which a practice workbook can't hold", "X024"))
        shared = []
        if "xl/sharedStrings.xml" in names:
            shared = [_text(si) for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(M + "si")]
        styles = _styles(z)
        books = {}
        for title, part in _sheet_parts(z):
            sh = Sheet(title)
            root = ET.fromstring(z.read(part))
            for c in root.iter(M + "c"):
                ref = c.get("r")
                r, col = _cell(ref)
                style = styles[int(c.get("s", "0"))] if styles else ("General", None, 0)
                code, fill, num_id = style
                if fill:
                    sh.fills[(r, col)] = fill
                if code != "General":
                    sh.formats[(r, col)] = code
                kind = c.get("t", "n")
                f = c.find(M + "f")
                v = c.find(M + "v")
                raw = v.text if v is not None else None
                where = f"{name} {title}!{ref}"
                if kind == "e" and reading:
                    sh.cells[(r, col)] = XlError(raw)
                    continue
                if kind == "e":
                    raise TemplateError(coded(f"{where} holds an error value ({raw}): a practice workbook can't", "X025"))
                if kind == "s":
                    value = shared[int(raw)] if raw is not None else None
                elif kind == "inlineStr":
                    is_ = c.find(M + "is")
                    value = _text(is_) if is_ is not None else None
                elif kind == "str":
                    value = raw if raw is not None else None
                elif kind == "b":
                    value = None if raw is None else raw == "1"
                else:
                    value = None
                    if raw not in (None, ""):
                        number = float(raw)
                        if number.is_integer() and abs(number) < 1e15:
                            number = int(number)
                        is_date = num_id in DATE_IDS or (num_id >= 164 and is_date_format(code))
                        value = _serial_to_model(float(raw)) if is_date else number
                if f is not None:
                    if f.get("t") in ("shared", "array", "dataTable") and not reading:
                        raise TemplateError(coded(f"{where}: a {f.get('t')} formula, which a practice workbook "
                                                  "can't hold: write the formula in each cell", "X026"))
                    sh.cells[(r, col)] = Formula("=" + (f.text or ""), saved=value)
                elif value is not None and value != "":
                    sh.cells[(r, col)] = value
            merged = root.find(M + "mergeCells")
            for m in (merged if merged is not None else []):
                a, b = m.get("ref").split(":")
                (r1, c1), (r2, c2) = _cell(a), _cell(b)
                sh.merged.append((r1, c1, r2, c2))
            books[title] = sh
    return books


DOCX_NAME = re.compile(r"[A-Za-z0-9_]{1,60}\.docx")


def load_docx_template(name):
    """The Word template data/templates/<name>.docx (a DOCUMENT's TEMPLATE, D115) -> its bytes."""
    if not DOCX_NAME.fullmatch(name or ""):
        raise TemplateError(coded(f"TEMPLATE {name!r}: a Word template is named like estimate_template.docx", "X028"))
    path = TEMPLATES / name
    if not path.is_file():
        raise TemplateError(coded(f"TEMPLATE {name}: there is no data/templates/{name}", "X029"))
    return path.read_bytes()


def load_template(name):
    """The template data/templates/<name>.xlsx -> {sheet name: Sheet}."""
    path = template_path(name)
    if not path.is_file():
        raise TemplateError(coded(f"template={name}: there is no data/templates/{name}.xlsx", "X027"))
    return load(path.read_bytes(), f"{name}.xlsx")
