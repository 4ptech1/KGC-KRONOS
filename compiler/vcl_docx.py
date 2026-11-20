"""Fill a Word template (VCL 0.7 part 3f, the DOCUMENT kind; docs/vep-0.7-documents.md, D115), with the standard
library only (zipfile and xml.etree), so the toolchain and KRONOS Offline's page need nothing new.

The template is marked up as data/templates/README.md says: {{Name}} fields anywhere (the body, tables, headers and
footers), a table row or column of fields that repeats, and sections between the paragraphs {{BEGIN Name}} and
{{END Name}}.

    template_fields(data)  -> what the template holds: fields per place, sections, tables (for the validator)
    fill(data, plan)       -> (the filled .docx bytes, a Report)

plan is a Plan: fields {name: text}; repeats [Repeat]; keeps {section: True or False}. Every text is already worked
out by the interpreter (vcl_interp.run_document). Filling keeps each run's formatting: a field is replaced inside the
run where its "{{" starts, after the paragraph's runs are read together (Word splits typing over runs); the rest of
the field's characters are taken out of the runs that held them. Pictures, styles, headers, footers and page fields
are kept as they are; every other part of the file is copied byte for byte.
"""
import copy
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

from errcodes import coded

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS
TAG = "{urn:kronos-template}t"   # a table's number in the template, while filling (taken out before writing)
FIELD = re.compile(r"\{\{([A-Za-z0-9_]+)\}\}")
MARKER = re.compile(r"^\{\{(BEGIN|END) ([A-Za-z0-9_]+)\}\}$")
PARTS = re.compile(r"word/(document|header\d+|footer\d+)\.xml$")
TEMPLATE_LIMIT = 2 * 1024 * 1024   # Boss's templates (plan 4c): up to 2 MB


class DocxError(Exception):
    """A template or a plan that can't be filled, in words for the person."""


@dataclass
class Repeat:
    kind: str                 # "row", "column" or "section"
    target: object            # a table number (row, column) or a section name
    rows: list                # [[text, ...], ...]: one repeat per row
    binds: dict               # field name -> column number (1-based) in a row
    match: tuple = None       # (column number, the section field it must be SAME as): a repeat inside a section


@dataclass
class Plan:
    fields: dict = field(default_factory=dict)
    repeats: list = field(default_factory=list)
    keeps: dict = field(default_factory=dict)


@dataclass
class Report:
    fields: dict = field(default_factory=dict)      # name -> the text written for it (FIELD_TEXT)
    sections: dict = field(default_factory=dict)    # name -> copies made (SECTION_COUNT)
    tables: list = field(default_factory=list)      # the filled body's tables: [[[cell text]]] (TABLE_*)
    text: str = ""                                  # every paragraph's text, one per line (TEXT_HAS)
    left_over: int = 0                              # "{{" left anywhere (LEFT_OVER)


# ---------------------------------------------------------------- reading

def _check(data):
    if len(data) > TEMPLATE_LIMIT:
        raise DocxError(coded(f"the Word template is {len(data) // 1024} KB; a template may be up to 2 MB", "X030"))
    if not zipfile.is_zipfile(io.BytesIO(data)):
        raise DocxError(coded("this is not a Word .docx file", "X031"))
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        if "word/document.xml" not in z.namelist():
            raise DocxError(coded("this is a zip file but not a Word .docx file", "X032"))


def _paragraph_text(p):
    return "".join(t.text or "" for t in p.iter(W + "t"))


def _tables(root):
    """The body's tables, outermost first, in document order (nested tables are not counted)."""
    body = root.find(W + "body")
    out = []

    def walk(e, inside):
        for child in e:
            if child.tag == W + "tbl":
                if not inside:
                    out.append(child)
                walk(child, True)
            else:
                walk(child, inside)
    walk(body, False)
    return out


def template_fields(data):
    """What a template holds: {"body": fields outside any repeat target, "header": ..., "tables": [[fields of each
    row]], "sections": {name: fields inside it}, "order": [section names]}, for the validator. Fields inside a
    section are listed under the section; a table's fields under its number (counted in the template's order)."""
    _check(data)
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        root = ET.fromstring(z.read("word/document.xml"))
        others = [ET.fromstring(z.read(n)) for n in sorted(z.namelist()) if PARTS.fullmatch(n) and "document" not in n]
    sections, order, current = {}, [], None
    body_fields = set()
    for p in root.iter(W + "p"):
        text = _paragraph_text(p)
        m = MARKER.match(text.strip())
        if m:
            if m.group(1) == "BEGIN":
                if current is not None:
                    raise DocxError(coded(f"section {m.group(2)} starts inside section {current}: sections can't nest", "X033"))
                if m.group(2) in sections:
                    raise DocxError(coded(f"the template has two sections named {m.group(2)}", "X034"))
                current = m.group(2)
                sections[current] = set()
                order.append(current)
            else:
                if current != m.group(2):
                    raise DocxError(coded(f"{{{{END {m.group(2)}}}}} doesn't close the section that is open "
                                          f"({current or 'none'})", "X035"))
                current = None
            continue
        if "{{BEGIN" in text or "{{END" in text:
            raise DocxError(coded(f"a section marker must be a paragraph on its own: {text.strip()!r}", "X036"))
        names = set(FIELD.findall(text))
        if current is not None:
            sections[current] |= names
        else:
            body_fields |= names
    if current is not None:
        raise DocxError(coded(f"section {current} has no {{{{END {current}}}}}", "X037"))
    tables = []
    for tbl in _tables(root):
        rows = [[set(FIELD.findall(_paragraph_text_all(tc))) for tc in tr.findall(W + "tc")]
                for tr in tbl.findall(W + "tr")]
        tables.append(rows)
    header = set()
    for part in others:
        for p in part.iter(W + "p"):
            header |= set(FIELD.findall(_paragraph_text(p)))
    table_section, current = [], None   # the section each table lies in (None: outside every section)
    for child in root.find(W + "body"):
        if child.tag == W + "p":
            m = MARKER.match(_paragraph_text(child).strip())
            if m:
                current = m.group(2) if m.group(1) == "BEGIN" else None
                continue
        holder = ET.Element("holder")
        holder.append(child)
        table_section += [current] * len(_tables_in(holder))
    return {"body": body_fields, "header": header, "tables": tables, "sections": sections, "order": order,
            "table_section": table_section}


def _tables_in(e):
    out = []

    def walk(x, inside):
        for child in x:
            if child.tag == W + "tbl":
                if not inside:
                    out.append(child)
                walk(child, True)
            else:
                walk(child, inside)
    walk(e, False)
    return out


def _paragraph_text_all(e):
    return "\n".join(_paragraph_text(p) for p in e.iter(W + "p"))


# ---------------------------------------------------------------- filling text inside runs

def _replace_in_paragraph(p, values):
    """Replace every {{Name}} of this paragraph whose name is in values; the text goes into the run where the
    field's "{{" starts, keeping that run's formatting. Returns the names replaced."""
    texts = [t for t in p.iter(W + "t")]
    if not texts:
        return []
    joined = "".join(t.text or "" for t in texts)
    found = [m for m in FIELD.finditer(joined) if m.group(1) in values]
    if not found:
        return []
    spans, pos = [], 0
    for t in texts:
        n = len(t.text or "")
        spans.append((pos, pos + n))
        pos += n
    new = [list(t.text or "") for t in texts]
    for m in reversed(found):
        start, end = m.span()
        first = True
        for i, (a, b) in enumerate(spans):
            lo, hi = max(a, start), min(b, end)
            if lo >= hi:
                continue
            piece = new[i]
            if first and lo == start:
                piece[lo - a:hi - a] = list(values[m.group(1)])
                first = False
            else:
                piece[lo - a:hi - a] = []
    for t, chars in zip(texts, new):
        t.text = "".join(chars)
        if t.text != t.text.strip() or "  " in t.text:
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return [m.group(1) for m in found]


def _fill_fields(e, values):
    done = []
    for p in list(e.iter(W + "p")):
        done += _replace_in_paragraph(p, values)
    return done


def _row_values(repeat, row):
    out = {}
    for name, col in repeat.binds.items():
        if not 1 <= col <= len(row):
            raise DocxError(coded(f"{name}={col}: the source row has {len(row)} column{'s' if len(row) != 1 else ''}", "X038"))
        out[name] = row[col - 1]
    return out


# ---------------------------------------------------------------- repeats

def _repeat_row(tbl, repeat, rows):
    names = set(repeat.binds)
    target = next((tr for tr in tbl.findall(W + "tr") if names & set(FIELD.findall(_paragraph_text_all(tr)))), None)
    if target is None:
        raise DocxError(coded(f"table {repeat.target} has no row with the fields {', '.join(sorted(names))}", "X039"))
    at = list(tbl).index(target)
    tbl.remove(target)
    for i, row in enumerate(rows):
        clone = copy.deepcopy(target)
        _fill_fields(clone, _row_values(repeat, row))
        tbl.insert(at + i, clone)


def _repeat_column(tbl, repeat, rows):
    names = set(repeat.binds)
    trs = tbl.findall(W + "tr")
    col = None
    for tr in trs:
        for i, tc in enumerate(tr.findall(W + "tc")):
            if names & set(FIELD.findall(_paragraph_text_all(tc))):
                col = i
                break
        if col is not None:
            break
    if col is None:
        raise DocxError(coded(f"table {repeat.target} has no column with the fields {', '.join(sorted(names))}", "X040"))
    grid = tbl.find(W + "tblGrid")
    gridcols = grid.findall(W + "gridCol") if grid is not None else []
    if grid is not None and col < len(gridcols):
        g = gridcols[col]
        at = list(grid).index(g)
        grid.remove(g)
        for i in range(len(rows)):
            grid.insert(at + i, copy.deepcopy(g))
    for tr in trs:
        tcs = tr.findall(W + "tc")
        if col >= len(tcs):
            continue
        cell = tcs[col]
        at = list(tr).index(cell)
        tr.remove(cell)
        for i, row in enumerate(rows):
            clone = copy.deepcopy(cell)
            _fill_fields(clone, _row_values(repeat, row))
            tr.insert(at + i, clone)


def _section_span(body, name):
    """(begin paragraph, end paragraph, the elements between) of a section among the body's own children."""
    kids = list(body)
    begin = end = None
    for i, k in enumerate(kids):
        if k.tag == W + "p":
            text = _paragraph_text(k).strip()
            if text == f"{{{{BEGIN {name}}}}}":
                begin = i
            elif text == f"{{{{END {name}}}}}":
                end = i
    if begin is None or end is None or end < begin:
        raise DocxError(coded(f"the template has no section {name} (paragraphs {{{{BEGIN {name}}}}} and "
                              f"{{{{END {name}}}}} in the body)", "X041"))
    return kids[begin], kids[end], kids[begin + 1:end]


def _tag_tables(root):
    for n, tbl in enumerate(_tables(root), start=1):
        tbl.set(TAG, str(n))


def _tagged(e, number):
    return [t for t in e.iter(W + "tbl") if t.get(TAG) == str(number)]


def _same(a, b):
    return str(a).strip().casefold() == str(b).strip().casefold()


def fill(data, plan):
    """The template filled as the plan says -> (bytes of the new .docx, Report)."""
    _check(data)
    report = Report()
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        names = z.namelist()
        raw = {n: z.read(n) for n in names}
        infos = {n: z.getinfo(n) for n in names}
    root = _parse(raw["word/document.xml"])
    body = root.find(W + "body")
    _tag_tables(root)
    sections = [r for r in plan.repeats if r.kind == "section"]
    inner = [r for r in plan.repeats if r.kind in ("row", "column") and r.match]
    outer = [r for r in plan.repeats if r.kind in ("row", "column") and not r.match]
    for r in sections:
        begin, end, middle = _section_span(body, r.target)
        at = list(body).index(begin)
        for e in [begin, *middle, end]:
            body.remove(e)
        made = []
        for i, row in enumerate(r.rows):
            values = _row_values(r, row)
            clones = [copy.deepcopy(e) for e in middle]
            holder = ET.Element("holder")
            holder.extend(clones)
            for nested in inner:   # a repeat inside this section: only the rows that match this section row
                col, key = nested.match
                if key not in values:
                    raise DocxError(coded(f"match={col}={key}: {key} is not a field of section {r.target}", "X042"))
                rows = [x for x in nested.rows if col <= len(x) and _same(x[col - 1], values[key])]
                for tbl in _tagged(holder, nested.target):
                    (_repeat_row if nested.kind == "row" else _repeat_column)(tbl, nested, rows)
            _fill_fields(holder, values)
            made += list(holder)
        for i, e in enumerate(made):
            body.insert(at + i, e)
        report.sections[r.target] = len(r.rows)
    for r in outer:
        tables = _tagged(root, r.target)
        if not tables:
            raise DocxError(coded(f"the template has no table {r.target}", "X043"))
        for tbl in tables:
            (_repeat_row if r.kind == "row" else _repeat_column)(tbl, r, r.rows)
    for name, keep in plan.keeps.items():
        begin, end, middle = _section_span(body, name)
        body.remove(begin)
        body.remove(end)
        if not keep:
            for e in middle:
                body.remove(e)
        report.sections[name] = 1 if keep else 0
    for name in _fill_fields(root, plan.fields):
        report.fields.setdefault(name, plan.fields[name])
    for t in root.iter(W + "tbl"):
        t.attrib.pop(TAG, None)
    out = {"word/document.xml": _write(root, raw["word/document.xml"])}
    texts = [_paragraph_text(p) for p in root.iter(W + "p")]
    for n in names:
        if PARTS.fullmatch(n) and n != "word/document.xml":
            part = _parse(raw[n])
            for name in _fill_fields(part, plan.fields):
                report.fields.setdefault(name, plan.fields[name])
            out[n] = _write(part, raw[n])
            texts += [_paragraph_text(p) for p in part.iter(W + "p")]
    report.text = "\n".join(texts)
    report.left_over = sum(t.count("{{") for t in texts)
    report.tables = [[[_paragraph_text_all(tc) for tc in tr.findall(W + "tc")] for tr in tbl.findall(W + "tr")]
                     for tbl in _tables(root)]
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n in names:
            info = zipfile.ZipInfo(n, infos[n].date_time)
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, out.get(n, raw[n]))
    return buf.getvalue(), report


# ---------------------------------------------------------------- XML in and out, namespaces kept

def _declarations(data):
    head = re.search(rb"<[A-Za-z_][\w.:-]*\s[^>]*>", data[data.find(b"<", data.find(b"?>") + 2 if b"?>" in data else 0):])
    return dict(re.findall(r'xmlns:([A-Za-z_][\w.-]*)="([^"]+)"', head.group(0).decode("utf-8"))) if head else {}


def _parse(data):
    for prefix, uri in _declarations(data).items():
        ET.register_namespace(prefix, uri)
    return ET.fromstring(data)


def _write(root, original):
    """The part's XML, with every namespace the original root declared (Word reads prefixes in mc:Ignorable)."""
    text = ET.tostring(root, encoding="unicode")
    declared = _declarations(original)
    m = re.match(r"<([A-Za-z_][\w.:-]*)", text)
    have = set(re.findall(r'xmlns:([A-Za-z_][\w.-]*)="', text[:text.find(">")]))
    extra = "".join(f' xmlns:{p}="{u}"' for p, u in declared.items() if p not in have)
    text = text[:m.end()] + extra + text[m.end():]
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + text).encode("utf-8")
