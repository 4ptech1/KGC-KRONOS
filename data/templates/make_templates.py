"""The look-alike KGC templates (Phase 2 of docs/kgc-documents-plan.md, D108): invented practice copies of KGC's own
documents, with the same layout as the forms the survey describes (docs/kgc-documents.md), for tests only.

    python data/templates/make_templates.py

writes, next to this file:
  wps_form.xlsx           a WPS form: merged cells, a logo and a weld sketch, a print area, a page break, three run rows,
                          and the heat-input check outside the print area
  wps_form_b.xlsx         a second WPS on the same form, with other numbers and PQR values (the old WPS a new one is
                          made from, D114)
  estimation.xlsx         a cost estimate in the same 7 sheets, with the same kinds of formulas
  enquiry_register.xlsx   an enquiry register: an enquiry's first columns merged over its item rows, formulas, notes
  estimate_template.docx  an internal estimation sheet with {{fields}}, a repeated row, a repeated column and a repeated
                          section, a header with a logo and a footer with page numbers
  proposal_template.docx  a technical proposal (Part A) with {{fields}}, repeated rows, a section repeated per item with
                          a repeated row inside it, and an optional section

Every name, number and price is invented. The workbooks hold each formula's value as Excel would have saved it, worked
out here, so a reader that takes "the last values saved by Excel" finds them. The files are the same on every run: fixed
zip dates and document properties. tests/test_templates.py checks that each committed file is what this script makes.
How a Word template is marked up is in README.md.
"""
import io
import math
import re
import struct
import sys
import zipfile
import zlib
from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).resolve().parent

import openpyxl  # noqa: E402
from openpyxl.comments import Comment  # noqa: E402
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side  # noqa: E402
from openpyxl.worksheet.pagebreak import Break  # noqa: E402

FILES = ("wps_form.xlsx", "wps_form_b.xlsx", "estimation.xlsx", "enquiry_register.xlsx", "estimate_template.docx",
         "proposal_template.docx")
ZIP_DATE = (2026, 11, 20, 0, 0, 0)
CORE = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>{title}</dc:title>'
        '<dc:creator>KRONOS look-alike template (invented data)</dc:creator>'
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-11-20T00:00:00Z</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-11-20T00:00:00Z</dcterms:modified></cp:coreProperties>')


# ---------------------------------------------------------------- pictures (a PNG made with the standard library)

def png(width, height, pixel):
    """A small RGB PNG; pixel(x, y) gives (r, g, b)."""
    raw = b"".join(b"\x00" + bytes(c for x in range(width) for c in pixel(x, y)) for y in range(height))

    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b""))


def logo_png():
    """An invented logo: a blue square with a white bar."""
    return png(48, 24, lambda x, y: (255, 255, 255) if 8 <= y < 16 and 6 <= x < 42 else (31, 78, 121))


def sketch_png():
    """An invented weld sketch: two grey plates and a double V between them."""
    def pixel(x, y):
        if abs(x - 40) <= abs(y - 20) // 2 + 1:
            return (200, 60, 40) if abs(x - 40) >= abs(y - 20) // 2 else (255, 255, 255)
        return (150, 150, 150) if 8 <= y < 32 else (255, 255, 255)
    return png(80, 40, pixel)


# ---------------------------------------------------------------- the workbooks

THIN = Side(style="thin")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOLD = Font(bold=True)
HEAD_FILL = PatternFill("solid", fgColor="DDEBF7")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
WRAP = Alignment(wrap_text=True, vertical="top")


class Book:
    """An openpyxl workbook plus each formula's saved value (sheet, cell) -> value, for after the save."""

    def __init__(self, title):
        self.wb = openpyxl.Workbook()
        self.title = title
        self.values = {}
        self.pictures = {}   # sheet title -> [(png bytes, from col, from row, to col, to row)] (0-based anchors)

    def sheet(self, name, first=False):
        if first:
            ws = self.wb.active
            ws.title = name
            return ws
        return self.wb.create_sheet(name)

    def put(self, ws, ref, value, fmt=None, bold=False, fill=None):
        c = ws[ref]
        c.value = value
        if fmt:
            c.number_format = fmt
        if bold:
            c.font = BOLD
        if fill:
            c.fill = fill
        return c

    def formula(self, ws, ref, text, value, fmt=None, bold=False):
        """A formula and the value Excel would save for it."""
        self.put(ws, ref, text, fmt, bold)
        self.values[(ws.title, ref)] = value

    def save(self):
        out = io.BytesIO()
        self.wb.save(out)
        return finish_xlsx(out.getvalue(), self)


def _cell_value_xml(value):
    if isinstance(value, bool):
        return ' t="b"', "1" if value else "0"
    if isinstance(value, str):
        return ' t="str"', escape(value)
    return "", repr(float(value)) if isinstance(value, float) else str(value)


def finish_xlsx(data, book):
    """Write each formula's saved value into the sheet XML, add the pictures, and fix the dates."""
    src = zipfile.ZipFile(io.BytesIO(data))
    sheets = sheet_parts(src)
    parts = {n: src.read(n) for n in src.namelist()}
    for title, part in sheets.items():
        xml = parts[part].decode("utf-8")
        for (t, ref), value in book.values.items():
            if t != title:
                continue
            pattern = re.compile(r'<c r="%s"((?: [a-z]+="[^"]*")*)><f>(.*?)</f><v\s*/?>(?:</v>)?</c>' % ref)
            kind, text = _cell_value_xml(value)
            xml, n = pattern.subn(lambda m: f'<c r="{ref}"{m.group(1)}{kind}><f>{m.group(2)}</f><v>{text}</v></c>', xml)
            if n != 1:
                raise RuntimeError(f"{title}!{ref}: formula cell not found in the saved XML")
        parts[part] = xml.encode("utf-8")
    for n, (title, pics) in enumerate(sorted(book.pictures.items()), 1):
        add_pictures(parts, sheets[title], n, pics)
    parts["docProps/core.xml"] = CORE.format(title=escape(book.title)).encode("utf-8")
    return write_zip(parts, list(parts))


def sheet_parts(z):
    """{sheet title: its part name} from workbook.xml and its rels."""
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    target = dict(re.findall(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    target.update({k: v for v, k in re.findall(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    out = {}
    for name, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb):
        t = target[rid].lstrip("/")
        out[name.replace("&amp;", "&")] = t if t.startswith("xl/") else "xl/" + t
    return out


def add_pictures(parts, sheet_part, n, pics):
    """Add one drawing part with the pictures to a sheet, as Excel stores them (a two-cell anchor each)."""
    drawing = f"xl/drawings/drawing{n}.xml"
    anchors, rels = [], []
    for i, (data, c1, r1, c2, r2) in enumerate(pics, 1):
        media = f"xl/media/template{n}_{i}.png"
        parts[media] = data
        rels.append(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                    f'relationships/image" Target="../media/{media.rsplit("/", 1)[1]}"/>')
        anchors.append(
            f'<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>{c1}</xdr:col><xdr:colOff>0</xdr:colOff>'
            f'<xdr:row>{r1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>{c2}</xdr:col>'
            f'<xdr:colOff>0</xdr:colOff><xdr:row>{r2}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
            f'<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{i + 1}" name="Picture {i}"/><xdr:cNvPicPr>'
            f'<a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill>'
            f'<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId{i}"/>'
            f'<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/>'
            f'<a:ext cx="0" cy="0"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>'
            f'<xdr:clientData/></xdr:twoCellAnchor>')
    parts[drawing] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' + "".join(anchors) + '</xdr:wsDr>'
    ).encode("utf-8")
    parts[f"xl/drawings/_rels/drawing{n}.xml.rels"] = rels_xml(rels)
    sheet_name = sheet_part.rsplit("/", 1)[1]
    rels_part = f"xl/worksheets/_rels/{sheet_name}.rels"
    rel = ('<Relationship Id="rIdTemplateDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
           f'relationships/drawing" Target="../drawings/drawing{n}.xml"/>')
    if rels_part in parts:
        parts[rels_part] = parts[rels_part].replace(b"</Relationships>", rel.encode("utf-8") + b"</Relationships>")
    else:
        parts[rels_part] = rels_xml([rel])
    xml = parts[sheet_part].decode("utf-8")
    tag = '<drawing r:id="rIdTemplateDrawing"/>'
    if 'xmlns:r=' not in xml.split(">", 2)[1]:
        xml = xml.replace("<worksheet ", '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/'
                          'relationships" ', 1)
    # the drawing goes after the page setup and breaks, before any legacy drawing (the schema's order)
    for before in ("<legacyDrawing", "<tableParts", "<extLst", "</worksheet>"):
        if before in xml:
            xml = xml.replace(before, tag + before, 1)
            break
    parts[sheet_part] = xml.encode("utf-8")
    types = parts["[Content_Types].xml"].decode("utf-8")
    if 'Extension="png"' not in types:
        types = types.replace("<Default ", '<Default Extension="png" ContentType="image/png"/><Default ', 1)
    types = types.replace("</Types>", f'<Override PartName="/{drawing}" ContentType="application/'
                          'vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
    parts["[Content_Types].xml"] = types.encode("utf-8")


def rels_xml(rels):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(rels) + "</Relationships>").encode("utf-8")


def write_zip(parts, order):
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name in order:
            info = zipfile.ZipInfo(name, ZIP_DATE)
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, parts[name])
    return out.getvalue()


def boxed(ws, area):
    for row in ws[area]:
        for c in row:
            c.border = BOX


# ---------------------------------------------------------------- the WPS form

K_FACTOR = 0.6
HI_MAX = 1.6
RUNS = [   # run, I min, I max, U min, U max, v min, v max, then the PQR's I min, I max, U min, U max, v, HI recorded
    ("1 (root)", 90, 130, 10, 13, 60, 90, 100, 120, 11, 12, 70, 0.74),
    ("2 (fill)", 140, 180, 12, 15, 80, 120, 150, 170, 13, 14, 125, 0.69),
    ("3 (cap)", 150, 190, 12, 16, 60, 130, 160, 180, 13, 15, 105, 0.93),
]


def heat_input(k, u, i, v):
    """Q = k x U x I x 60 / (v x 1000), kJ/mm: the WPS form's own formula."""
    return k * u * i * 60 / (v * 1000)


def excel_text_2dp(x):
    """TEXT(x, "0.00"): two decimals, a half rounded away from zero."""
    return f"{math.floor(x * 100 + 0.5) / 100:.2f}"


RUNS_B = [   # a second WPS on the same form, as KGC's two samples are (D114): other PQR values, every run in range
    ("1 (root)", 80, 120, 10, 12, 60, 90, 90, 110, 10, 12, 75, 0.58),
    ("2 (fill)", 130, 170, 11, 14, 70, 110, 140, 160, 12, 13, 95, 0.79),
    ("3 (cap)", 140, 180, 12, 15, 80, 120, 150, 170, 13, 14, 100, 0.86),
]
WPS_IDS = {"a": ("DM-WPS-201", "DM-PQR-201", "PQR-201", RUNS), "b": ("DM-WPS-105", "DM-PQR-105", "PQR-105", RUNS_B)}


def wps_form_b():
    return wps_form("b")


def wps_form(variant="a"):
    wps_no, pqr_no, pqr_short, runs = WPS_IDS[variant]
    b = Book("WPS form (look-alike)" if variant == "a" else "WPS form, second WPS (look-alike)")
    ws = b.sheet("WPS", first=True)
    for letter, width in zip("ABCDEFGHI", (24, 14, 14, 12, 12, 14, 14, 12, 12)):
        ws.column_dimensions[letter].width = width
    for page_top in (1, 34):
        ws.merge_cells(f"A{page_top}:B{page_top + 2}")
        ws.merge_cells(f"C{page_top}:G{page_top + 1}")
        ws.merge_cells(f"C{page_top + 2}:G{page_top + 2}")
        for r in range(page_top, page_top + 3):
            ws.merge_cells(f"H{r}:I{r}")
        b.put(ws, f"C{page_top}", "Welding Procedure Specification (WPS)", bold=True)
        b.put(ws, f"H{page_top}", "FORMAT NO:- QF/DM/31")
        b.put(ws, f"H{page_top + 1}", "DATE:- 02/11/2026")
        b.put(ws, f"C{page_top + 2}", "ISO 15609-1 : 2019")
        b.put(ws, f"H{page_top + 2}", "WPS Revision : R0")

    def row(r, label, actual=None, qualified=None, wide=False):
        ws.merge_cells(f"A{r}:B{r}")
        b.put(ws, f"A{r}", label, bold=True)
        if wide:
            ws.merge_cells(f"C{r}:I{r}")
            b.put(ws, f"C{r}", actual)
        else:
            ws.merge_cells(f"C{r}:E{r}")
            ws.merge_cells(f"F{r}:I{r}")
            b.put(ws, f"C{r}", actual)
            b.put(ws, f"F{r}", qualified)

    row(4, "Company Name :", "Demo Fabricators Pvt Ltd", wide=True)
    row(5, "Company Address :", "Plot 12, Example Industrial Estate", wide=True)
    row(6, "Welding Procedure Specification No. :", wps_no, "Rev. R0 - first issue")
    row(7, "WPQR No. :", pqr_no, "Certificate DM/2026/014")
    row(8, "Welding Process :", "GTAW (141) - single process", wide=True)
    row(9, "Degree of mechanization :", "Manual", wide=True)
    row(10, "Mode of Metal Transfer :", "Not applicable", wide=True)
    ws.merge_cells("A11:I11")
    b.put(ws, "A11", "Weld Preparation Details (Sketch)", bold=True)
    row(12, "Method of preparation and cleaning", "Machining; clean 25 mm each side", wide=True)
    row(13, "Welding Position", "Actual: PF", "Qualified: PA, PB, PF")
    ws.merge_cells("A14:C19")    # the sketch sits here
    ws.merge_cells("D14:I19")
    b.put(ws, "D14", "Double V butt joint; root gap 2-3 mm; root face 1 mm; back-gouge side B.")
    ws["D14"].alignment = WRAP
    row(20, "Weld Details", "Actual: ml bs", "Qualified: ml bs, ml ss nb")
    row(21, "Parent Material", "Actual", "Qualified")
    for r, label, actual, qualified in (
            (22, "Parent Material Type (Plate or Pipe)", "Plate", "Plate, and pipe OD > 500 mm"),
            (23, "Parent Material Designation", "SA 516 Gr 70", "Group 1 steels"),
            (24, "Parent Material Group (ISO/TR 15608)", "1.1", "1.1"),
            (25, "Material Thickness (mm)", 12, "6 to 24 mm"),
            (26, "Outside Pipe Diameter (mm)", "NA", "NA")):
        ws.merge_cells(f"A{r}:B{r}")
        ws.merge_cells(f"F{r}:I{r}")
        b.put(ws, f"A{r}", label, bold=True)
        b.put(ws, f"C{r}", actual)
        b.put(ws, f"D{r}", "#")
        b.put(ws, f"E{r}", actual)
        b.put(ws, f"F{r}", qualified)
    ws.merge_cells("A27:I27")
    b.put(ws, "A27", "Welding Details", bold=True)
    ws.merge_cells("H28:I28")
    for col, head in zip("ABCDEFGH", ("Run", "Welding Process", "Size of Filler Material (mm)", "Current (A)",
                                      "Voltage (V)", "Type of Current / Polarity", "Travel Speed (mm/min)",
                                      "Heat Input (kJ/mm) / calculated")):
        b.put(ws, f"{col}28", head, bold=True, fill=HEAD_FILL)
        ws[f"{col}28"].alignment = WRAP

    # the heat-input check, outside the print area
    ws.merge_cells("K24:AC24")
    b.put(ws, "K24", "HEAT-INPUT CHECK - outside the print area (not printed)", bold=True)
    ws.merge_cells("K25:M25")
    ws.merge_cells("K26:M26")
    ws.merge_cells("O25:AC26")
    b.put(ws, "K25", "k-factor, process 141")
    b.put(ws, "N25", K_FACTOR, fill=INPUT_FILL)
    b.put(ws, "O25", "Heat input Q = k x U x I x 60 / (v x 1000), kJ/mm")
    b.put(ws, "K26", "WPS max heat input, kJ/mm")
    b.put(ws, "N26", HI_MAX, fill=INPUT_FILL)
    for col, head in zip(("K L M N O P Q R S T V W X Y Z AA AB AC".split()),
                         ("Run", "I min (A)", "I max (A)", "U min (V)", "U max (V)", "v min (mm/min)",
                          "v max (mm/min)", "HI min (kJ/mm)", "HI max (kJ/mm)", "Max <= limit?", "PQR I min",
                          "PQR I max", "PQR U min", "PQR U max", "PQR v", "PQR HI recorded",
                          "PQR HI calc. (max A, V)", "PQR run inside WPS range?")):
        b.put(ws, f"{col}28", head, bold=True, fill=HEAD_FILL)

    for n, (run, i1, i2, u1, u2, v1, v2, pi1, pi2, pu1, pu2, pv, phi) in enumerate(runs):
        r = 29 + n
        ws.merge_cells(f"H{r}:I{r}")
        b.put(ws, f"A{r}", run)
        b.put(ws, f"B{r}", "GTAW (141)")
        b.put(ws, f"C{r}", "2.4")
        b.put(ws, f"F{r}", "DCEN")
        b.put(ws, f"K{r}", f"Run {run}")
        for col, v in zip("LMNOPQ", (i1, i2, u1, u2, v1, v2)):
            b.put(ws, f"{col}{r}", v, fill=INPUT_FILL)
        for col, v in zip(("V", "W", "X", "Y", "Z", "AA"), (pi1, pi2, pu1, pu2, pv, phi)):
            b.put(ws, f"{col}{r}", v)
        lo = heat_input(K_FACTOR, u1, i1, v2)
        hi = heat_input(K_FACTOR, u2, i2, v1)
        b.formula(ws, f"D{r}", f'=L{r}&"–"&M{r}', f"{i1}–{i2}")
        b.formula(ws, f"E{r}", f'=N{r}&"–"&O{r}', f"{u1}–{u2}")
        b.formula(ws, f"G{r}", f'=P{r}&"–"&Q{r}', f"{v1}–{v2}")
        b.formula(ws, f"H{r}", f'=TEXT(R{r},"0.00")&" – "&TEXT(S{r},"0.00")',
                  f"{excel_text_2dp(lo)} – {excel_text_2dp(hi)}")
        b.formula(ws, f"R{r}", f"=$N$25*N{r}*L{r}*60/(Q{r}*1000)", lo, "0.000")
        b.formula(ws, f"S{r}", f"=$N$25*O{r}*M{r}*60/(P{r}*1000)", hi, "0.000")
        b.formula(ws, f"T{r}", f'=IF(S{r}<=$N$26,"OK","OVER LIMIT")', "OK" if hi <= HI_MAX else "OVER LIMIT")
        b.formula(ws, f"AB{r}", f"=$N$25*Y{r}*W{r}*60/(Z{r}*1000)", heat_input(K_FACTOR, pu2, pi2, pv), "0.000")
        inside = pi1 >= i1 and pi2 <= i2 and pu1 >= u1 and pu2 <= u2 and v1 <= pv <= v2
        b.formula(ws, f"AC{r}", f'=IF(AND(V{r}>=L{r},W{r}<=M{r},X{r}>=N{r},Y{r}<=O{r},Z{r}>=P{r},Z{r}<=Q{r}),'
                  f'"Yes","No")', "Yes" if inside else "No")
    boxed(ws, "A28:I31")
    ws.merge_cells("A32:I32")
    b.put(ws, "A32", "Additional runs as required within the ranges above.")
    ws.merge_cells("A33:I33")
    b.put(ws, "A33", "Continued to Page No. 2 ....")
    ws.row_breaks.append(Break(id=33))

    # page 2
    ws.merge_cells("C37:E37")
    ws.merge_cells("F37:I37")
    b.put(ws, "C37", f"Actual ({pqr_short})", bold=True)
    b.put(ws, "F37", "Qualified / WPS range", bold=True)
    for r, label, actual, qualified in (
            (38, "Weld Deposited Thickness", "GTAW (141) = 12 mm", "Max. 24 mm"),
            (39, "Filler Material Designation :", "AWS A5.18 ER70S-2", "AWS A5.18 ER70S-2"),
            (40, "Shielding Gas", "ISO 14175 - I1 (Ar)", "ISO 14175 - I1 (Ar)"),
            (41, "Gas Flow Rate (l/min)", "12-15", "10-18"),
            (42, "Preheat Temperature (°C)", "15 °C min.", "15 °C min."),
            (43, "Interpass Temperature (°C)", "250 °C max.", "250 °C max."),
            (44, "Post Weld Heat Treatment", "None", "Not required"),
            (45, "Heat Input (kJ/mm)", "0.69-0.93 recorded", "Max. 1.60 for any run")):
        row(r, label, actual, qualified)
    ws.merge_cells("B46:I46")
    ws.merge_cells("B47:I47")
    b.put(ws, "A46", "Remarks :", bold=True)
    b.put(ws, "B46", "1. Clean the joint before each run.")
    b.put(ws, "B47", "2. Record current and voltage for every run.")
    ws.merge_cells("B48:E48")
    ws.merge_cells("F48:I48")
    b.put(ws, "B48", "Prepared By / Reviewed & Approved", bold=True)
    b.put(ws, "F48", "Examiner / Examining Body :", bold=True)
    for r, label in ((49, "Name >>"), (50, "Date >>")):
        ws.merge_cells(f"B{r}:E{r}")
        ws.merge_cells(f"F{r}:I{r}")
        b.put(ws, f"A{r}", label, bold=True)
    b.put(ws, "B49", "Welding Engineer (demo)")
    b.put(ws, "F49", "Demo Inspection Body")
    b.put(ws, "B50", "02/11/2026")
    ws.print_area = "A1:I50"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    b.pictures["WPS"] = [(logo_png(), 0, 0, 2, 3), (sketch_png(), 0, 13, 3, 19)]
    return b.save()


# ---------------------------------------------------------------- the estimation workbook

INPUTS = [   # ref, parameter, value, unit, basis
    ("fx", "Exchange rate", 85, "INR/USD", "Invented rate for practice"),
    ("dens", "Steel density", 7.85e-06, "kg/mm3", "Carbon / low-alloy steel"),
    ("r_plate", "Plate SA 516 Gr 70, practice rate", 95, "INR/kg", "Invented"),
    ("r_forg", "Forged ring SA 105, practice rate", 130, "INR/kg", "Invented"),
    ("r_lab", "Fabrication labour rate", 650, "INR/hr", "Invented"),
    ("pm", "Plate cutting margin per edge", 20, "mm", "Invented"),
    ("scrap", "Plate nesting / scrap allowance", 0.08, "-", "Invented"),
    ("fa", "Forging machining allowance per surface", 5, "mm", "Invented"),
    ("oh", "Factory overheads", 0.1, "-", "Invented"),
    ("cont", "Contingency", 0.02, "-", "Invented"),
    ("margin", "Margin", 0.12, "-", "Invented"),
]
BOM = [   # scope, item, description, material, form, qty/set, sets, L or OD, W or ID, T or H, BOM finished wt (2 sets)
    ("CH", "101", "Channel flange", "SA 105", "Ring", 1, 2, 900, 700, 120, 640),
    ("CH", "102", "Channel shell", "SA 516 Gr 70", "Plate", 1, 2, 2300, 500, 16, 290),
    ("FH", "201", "Floating head flange", "SA 105", "Ring", 1, 2, 820, 640, 110, 470),
    ("FH", "202", "Floating head dish", "SA 516 Gr 70", "Plate", 1, 2, 900, 900, 18, 220),
]
CONVERSION = [   # sl, scope, activity, basis, qty/set, sets, unit, rate
    (1, "ALL", "Engineering & documentation", "Design check, drawings, QAP", 1, 1, "Lot", 120000),
    (2, "CH", "Channel fabrication", "Fit-up, welding, grinding", 60, 2, "hr", 650),
    (3, "FH", "Floating head fabrication", "Fit-up, welding, grinding", 45, 2, "hr", 650),
    (4, "ALL", "Hydrotest & painting", "Shop hydrotest, one coat primer", 1, 2, "Lot", 15000),
]
HYDRO = [   # option, scope, item, qty, rate
    (1, "CH", "Extra NDT - channel", 2, 18000),
    (1, "FH", "Extra NDT - floating head", 2, 14000),
    (2, "CH", "Test cover - channel", 1, 60000),
    (2, "FH", "Test ring - floating head", 1, 45000),
]
REGISTER_HEADS = ["Sr.Nos.", "Enquiry", "Date", "Result", "End User", "EPC /PMC", "LSTK", "Eng Contractor",
                  "Project", "Item", "Type", "U Stamp", "Old Tag No", "Revised Tag No", "Qty", "MOC", "Shell OD",
                  "Shell ID mm", "Shell THK mm", "TL To TL mm", "Drg Unit Wt Kgs", "Actual Wt", "Design Pr Kg/cm2",
                  "Design Temp °C", "Sour, NACE", "HIC/SSC Testing", "PWHT"]
ASSUMPTIONS = ["Prices ex-works, packing included; freight and taxes extra.",
               "Rates are invented practice rates.",
               "Option 1 has no shop hydrotest; option 2 uses a test cover and ring.",
               "Quantities are per the practice enquiry."]


def excel_roundup(x, digits):
    """ROUNDUP(x, digits): away from zero, to 10^-digits."""
    f = 10.0 ** -digits
    return math.copysign(math.ceil(abs(x) / f - 1e-9) * f, x)


def excel_round(x, digits):
    """ROUND(x, digits): a half rounded away from zero."""
    f = 10.0 ** digits
    return math.copysign(math.floor(abs(x) * f + 0.5) / f, x)


def estimation():
    b = Book("Estimation workbook (look-alike)")
    summary = b.sheet("Summary", first=True)
    inputs = b.sheet("Inputs")
    bom = b.sheet("BOM-RM")
    conv = b.sheet("Conversion")
    hydro = b.sheet("Hydrotest Options")
    enq = b.sheet("Enq Summary")
    assume = b.sheet("Assumptions & Clarif.")
    inp = {}   # ref -> (cell, value)

    b.put(inputs, "A1", "Estimation Inputs - Rates & Factors (invented, for practice)", bold=True)
    for col, head in zip("ABCDE", ("Ref", "Parameter", "Value", "Unit", "Basis / Remarks")):
        b.put(inputs, f"{col}4", head, bold=True, fill=HEAD_FILL)
    for n, (ref, name, value, unit, basis) in enumerate(INPUTS):
        r = 5 + n
        for col, v in zip("ABCDE", (ref, name, value, unit, basis)):
            b.put(inputs, f"{col}{r}", v, fill=INPUT_FILL if col == "C" else None)
        inp[ref] = (f"Inputs!$C${r}", value)

    def I(ref):
        return inp[ref][0]

    def V(ref):
        return inp[ref][1]

    b.put(bom, "A1", "Bill of Materials - Raw Material (2 sets)", bold=True)
    for col, head in zip("ABCDEFGHIJKLMNOPQR", (
            "Scope", "BOM Item", "Description", "Material", "Form", "Qty / set", "Sets", "Total Qty",
            "Fin. L / OD (mm)", "Fin. W / ID (mm)", "Fin. T / H (mm)", "BOM Fin. Wt, 2 sets (kg)", "Raw Wt each (kg)",
            "Raw Wt total (kg)", "Rate", "Rate unit", "Amount (INR)", "Remarks")):
        b.put(bom, f"{col}4", head, bold=True, fill=HEAD_FILL)
    totals = {"CH": [0, 0, 0], "FH": [0, 0, 0]}
    for n, (scope, item, desc, mat, form, qty, sets, d1, d2, d3, fin) in enumerate(BOM):
        r = 5 + n
        for col, v in zip("ABCDEFGIJKL", (scope, item, desc, mat, form, qty, sets, d1, d2, d3, fin)):
            b.put(bom, f"{col}{r}", v)
        b.formula(bom, f"H{r}", f"=F{r}*G{r}", qty * sets)
        if form == "Ring":
            fa = V("fa")
            each = math.pi / 4 * ((d1 + 2 * fa) ** 2 - (d2 - 2 * fa) ** 2) * (d3 + 2 * fa) * V("dens")
            b.formula(bom, f"M{r}", f"=PI()/4*((I{r}+2*{I('fa')})^2-(J{r}-2*{I('fa')})^2)*(K{r}+2*{I('fa')})"
                      f"*{I('dens')}", each, "0.0")
            rate_ref = "r_forg"
        else:
            pm = V("pm")
            each = (d1 + 2 * pm) * (d2 + 2 * pm) * d3 * V("dens") * (1 + V("scrap"))
            b.formula(bom, f"M{r}", f"=(I{r}+2*{I('pm')})*(J{r}+2*{I('pm')})*K{r}*{I('dens')}*(1+{I('scrap')})",
                      each, "0.0")
            rate_ref = "r_plate"
        total_wt = each * qty * sets
        amount = total_wt * V(rate_ref)
        b.formula(bom, f"N{r}", f"=M{r}*H{r}", total_wt, "0.0")
        b.formula(bom, f"O{r}", f"={I(rate_ref)}", V(rate_ref))
        b.put(bom, f"P{r}", "INR/kg")
        b.formula(bom, f"Q{r}", f"=N{r}*O{r}", amount, "#,##0")
        for k, v in enumerate((fin, total_wt, amount)):
            totals[scope][k] += v
    last = 4 + len(BOM)
    for r, label, scope in ((10, "Total - Channel Heads (CH)", "CH"), (11, "Total - Floating Heads (FH)", "FH")):
        b.put(bom, f"C{r}", label, bold=True)
        for k, col in enumerate("LNQ"):
            b.formula(bom, f"{col}{r}", f'=SUMIF($A$5:$A${last},"{scope}",{col}5:{col}{last})', totals[scope][k],
                      "#,##0")
    b.put(bom, "C12", "TOTAL RAW MATERIAL (2 sets)", bold=True)
    for k, col in enumerate("LNQ"):
        b.formula(bom, f"{col}12", f"={col}10+{col}11", totals["CH"][k] + totals["FH"][k], "#,##0")

    b.put(conv, "A1", "Conversion / Fabrication Cost Estimate - 2 Sets", bold=True)
    for col, head in zip("ABCDEFGHIJ", ("Sl", "Scope", "Activity", "Basis / Remarks", "Qty / set", "Sets",
                                        "Total Qty", "Unit", "Rate (INR)", "Amount (INR)")):
        b.put(conv, f"{col}4", head, bold=True, fill=HEAD_FILL)
    conv_tot = {"CH": 0, "FH": 0, "ALL": 0}
    for n, (sl, scope, act, basis, qty, sets, unit, rate) in enumerate(CONVERSION):
        r = 5 + n
        for col, v in zip("ABCDEFHI", (sl, scope, act, basis, qty, sets, unit, rate)):
            b.put(conv, f"{col}{r}", v)
        b.formula(conv, f"G{r}", f"=E{r}*F{r}", qty * sets)
        b.formula(conv, f"J{r}", f"=G{r}*I{r}", qty * sets * rate, "#,##0")
        conv_tot[scope] += qty * sets * rate
    last = 4 + len(CONVERSION)
    for r, scope in ((10, "CH"), (11, "FH"), (12, "ALL")):
        b.put(conv, f"I{r}", f"Total {scope}", bold=True)
        b.formula(conv, f"J{r}", f'=SUMIF($B$5:$B${last},"{scope}",J5:J{last})', conv_tot[scope], "#,##0")
    b.formula(conv, "J13", "=SUM(J10:J12)", sum(conv_tot.values()), "#,##0")

    b.put(hydro, "A1", "Hydrotest Options - Cost Adders (2 Sets)", bold=True)
    for col, head in zip("ABCEKL", ("Opt", "Scope", "Item", "Qty", "Rate (INR)", "Amount (INR)")):
        b.put(hydro, f"{col}4", head, bold=True, fill=HEAD_FILL)
    adders = {}
    for n, (opt, scope, item, qty, rate) in enumerate(HYDRO):
        r = 5 + n
        for col, v in zip("ABCEK", (opt, scope, item, qty, rate)):
            b.put(hydro, f"{col}{r}", v)
        b.formula(hydro, f"L{r}", f"=E{r}*K{r}", qty * rate, "#,##0")
        adders[(opt, scope)] = adders.get((opt, scope), 0) + qty * rate
    last = 4 + len(HYDRO)
    for r, (opt, scope) in zip((10, 11, 12, 13), ((1, "CH"), (1, "FH"), (2, "CH"), (2, "FH"))):
        b.put(hydro, f"K{r}", f"Option {opt} {scope}", bold=True)
        b.formula(hydro, f"L{r}", f'=SUMIFS($L$5:$L${last},$A$5:$A${last},"{opt}",$B$5:$B${last},"{scope}")',
                  adders[(opt, scope)], "#,##0")

    s = summary
    b.put(s, "A1", "COST ESTIMATE SUMMARY - Practice Channel & Floating Head Assemblies, DM-1001A/B", bold=True)
    b.put(s, "A2", "End User: Example Refinery Ltd (invented)  |  OEM: Demo OEM, practice drawing")
    for col, head in zip("ABCD", ("Cost Build-up (INR, 2 sets)", "Channel Heads (2)", "Floating Heads (2)", "Total")):
        b.put(s, f"{col}4", head, bold=True, fill=HEAD_FILL)
    v = {}

    def line(r, label, b_text, b_val, c_text, c_val, fmt="#,##0"):
        b.put(s, f"A{r}", label)
        b.formula(s, f"B{r}", b_text, b_val, fmt)
        b.formula(s, f"C{r}", c_text, c_val, fmt)
        b.formula(s, f"D{r}", f"=B{r}+C{r}", b_val + c_val, fmt)
        v[r] = (b_val, c_val)

    line(5, "Raw material (BOM-RM)", "='BOM-RM'!$Q$10", totals["CH"][2], "='BOM-RM'!$Q$11", totals["FH"][2])
    line(6, "Conversion - direct", "=Conversion!$J$10", conv_tot["CH"], "=Conversion!$J$11", conv_tot["FH"])
    line(7, "Conversion - common (50:50)", "=Conversion!$J$12/2", conv_tot["ALL"] / 2, "=Conversion!$J$12/2",
         conv_tot["ALL"] / 2)
    base = (sum(v[r][0] for r in (5, 6, 7)), sum(v[r][1] for r in (5, 6, 7)))
    b.put(s, "A8", "Base factory cost")
    b.formula(s, "B8", "=SUM(B5:B7)", base[0], "#,##0")
    b.formula(s, "C8", "=SUM(C5:C7)", base[1], "#,##0")
    b.formula(s, "D8", "=SUM(D5:D7)", base[0] + base[1], "#,##0")
    line(9, "Overheads", f"=B8*{I('oh')}", base[0] * V("oh"), f"=C8*{I('oh')}", base[1] * V("oh"))
    line(10, "Contingency", f"=B8*{I('cont')}", base[0] * V("cont"), f"=C8*{I('cont')}", base[1] * V("cont"))
    total = tuple(base[k] + v[9][k] + v[10][k] for k in (0, 1))
    line(11, "Total base cost", "=B8+B9+B10", total[0], "=C8+C9+C10", total[1])
    oh = 1 + V("oh")
    line(12, "Option 1 adder (extra NDT)", f"='Hydrotest Options'!$L$10*(1+{I('oh')})", adders[(1, "CH")] * oh,
         f"='Hydrotest Options'!$L$11*(1+{I('oh')})", adders[(1, "FH")] * oh)
    line(13, "Option 2 adder (test cover and ring)", f"='Hydrotest Options'!$L$12*(1+{I('oh')})",
         adders[(2, "CH")] * oh, f"='Hydrotest Options'!$L$13*(1+{I('oh')})", adders[(2, "FH")] * oh)
    line(14, "Total cost - Option 1", "=B11+B12", total[0] + v[12][0], "=C11+C12", total[1] + v[12][1])
    line(15, "Total cost - Option 2", "=B11+B13", total[0] + v[13][0], "=C11+C13", total[1] + v[13][1])
    for col, head in zip("ABCD", ("SELLING PRICE (Ex-Works)", "Channel Head (per No.)", "Floating Head (per No.)",
                                  "Total 2 Sets")):
        b.put(s, f"{col}16", head, bold=True, fill=HEAD_FILL)
    m = 1 + V("margin")
    prices = {}
    for r, label, cost_row in ((17, "Option 1 (INR)", 14), (18, "Option 2 (INR)", 15)):
        b.put(s, f"A{r}", label)
        pb = excel_roundup(v[cost_row][0] * m / 2, -3)
        pc = excel_roundup(v[cost_row][1] * m / 2, -3)
        b.formula(s, f"B{r}", f"=ROUNDUP(B{cost_row}*(1+{I('margin')})/2,-3)", pb, "#,##0")
        b.formula(s, f"C{r}", f"=ROUNDUP(C{cost_row}*(1+{I('margin')})/2,-3)", pc, "#,##0")
        b.formula(s, f"D{r}", f"=2*B{r}+2*C{r}", 2 * pb + 2 * pc, "#,##0")
        prices[r] = (pb, pc, 2 * pb + 2 * pc)
    b.put(s, "A19", "Option 1 (USD)")
    for k, col in enumerate("BCD"):
        b.formula(s, f"{col}19", f"={col}17/{I('fx')}", prices[17][k] / V("fx"), "#,##0")
    b.put(s, "A20", "Selling price per kg finished - Opt 1 (INR/kg)")
    fin = totals["CH"][0] + totals["FH"][0]
    b.formula(s, "B20", "=D17/('BOM-RM'!$L$10+'BOM-RM'!$L$11)", prices[17][2] / fin, "#,##0")
    b.put(s, "F4", "Notes", bold=True)
    b.put(s, "F5", "1. Invented practice estimate: no real job, client or price.")

    heads = REGISTER_HEADS + ["Unit Price Opt-1 (INR)", "Unit Price Opt-2 (INR)"]
    for n, head in enumerate(heads, 1):
        c = enq.cell(1, n, head)
        c.font = BOLD
        c.fill = HEAD_FILL
    for r, (item, tag, idx, fin_ref, price_col) in enumerate(
            (("Channel Head Assembly", "DM-1001A", 0, "$L$10", "B"),
             ("Floating Head Assembly", "DM-1001B", 1, "$L$11", "C")), 2):
        row = [r - 1, "E-9100 (TBA)", "02.11.2026", None, "Example Refinery Ltd", "Demo EPC", None, "Demo OEM",
               "Practice Expansion", item, "HE Part", "No", tag, tag, 2, "SA 516 Gr 70"]
        for n, value in enumerate(row, 1):
            if value is not None:
                enq.cell(r, n, value)
        shell_id, thk = (700, 16) if idx == 0 else (640, 18)
        enq[f"R{r}"] = shell_id
        enq[f"S{r}"] = thk
        enq[f"T{r}"] = 1200 if idx == 0 else 900
        b.formula(enq, f"Q{r}", f"=R{r}+2*S{r}", shell_id + 2 * thk)
        b.formula(enq, f"U{r}", f"='BOM-RM'!{fin_ref}/2", totals["CH" if idx == 0 else "FH"][0] / 2, "0.0")
        b.formula(enq, f"W{r}", "=ROUND(20*1.01972,2)", excel_round(20 * 1.01972, 2))
        enq[f"X{r}"] = 250
        enq[f"Y{r}"] = "No"
        enq[f"Z{r}"] = "No"
        enq[f"AA{r}"] = "No"
        b.formula(enq, f"AB{r}", f"=Summary!{price_col}17", prices[17][idx], "#,##0")
        b.formula(enq, f"AC{r}", f"=Summary!{price_col}18", prices[18][idx], "#,##0")

    b.put(assume, "A1", "Assumptions & Clarifications", bold=True)
    b.put(assume, "A3", "No.", bold=True, fill=HEAD_FILL)
    b.put(assume, "B3", "Assumption / clarification", bold=True, fill=HEAD_FILL)
    for n, text in enumerate(ASSUMPTIONS, 1):
        b.put(assume, f"A{3 + n}", n)
        b.put(assume, f"B{3 + n}", text)
    return b.save()


# ---------------------------------------------------------------- the enquiry register

ENQUIRIES = [   # sr, enquiry, date, result, end user, epc, lstk, contractor, project, items
    (1, "E-9101", "05.10.2026", "Lost", "Example Gas Co", "Demo EPC", None, "Demo Engineers", "PRJ-101", [
        ("1st Stage Suction Bottle", "PV", "Yes", "DM-V-101A/B", 2, "SA 516 Gr 70N", 400, 10, 1600, 410, 18, 120),
        ("1st Stage Discharge Bottle", "PV", "Yes", "DM-V-102A/B", 2, "SA 516 Gr 70N", 450, 12, 1700, 470, 35, 180)]),
    (2, "E-9102", "19.10.2026", None, "Sample Oils Ltd", "Practice EPC", "Demo LSTK", "Practice EPC", "PRJ-205", [
        ("Knock-out Drum", "PV", "Yes", "TBA", 1, "SA 516 Gr 70", 900, 14, 2400, 1250, 12, 90),
        ("Suction Bottle", "PV", "Yes", "TBA", 2, "SA 516 Gr 70", 500, 12, 1500, 520, 25, 150),
        ("Discharge Bottle", "PV", "Yes", "TBA", 2, "SA 516 Gr 70", 500, 16, 1500, 610, 60, 150)]),
    (3, "E-9103", "02.11.2026", None, "Example Refinery Ltd", None, None, "Demo OEM", "PRJ-310", [
        ("Channel Head Assembly", "HE Part", "No", "DM-1001A", 1, "SA 516 Gr 70", 700, 16, 1200, 360, 20, 250)]),
]


def enquiry_register():
    b = Book("Enquiry register (look-alike)")
    ws = b.sheet("Sheet1", first=True)
    for n, head in enumerate(REGISTER_HEADS, 1):
        c = ws.cell(1, n, head)
        c.font = BOLD
        c.fill = HEAD_FILL
        c.alignment = WRAP
    r = 2
    for sr, enquiry, date, result, end_user, epc, lstk, contractor, project, items in ENQUIRIES:
        first = r
        for value, col in zip((sr, enquiry, date, result, end_user, epc, lstk, contractor, project), "ABCDEFGHI"):
            if value is not None:
                ws[f"{col}{first}"] = value
        for item, kind, stamp, tag, qty, moc, shell_id, thk, tl, wt, barg, temp in items:
            for col, value in zip("JKLNOPRSTUX", (item, kind, stamp, tag, qty, moc, shell_id, thk, tl, wt, temp)):
                ws[f"{col}{r}"] = value
            b.formula(ws, f"Q{r}", f"=R{r}+2*S{r}", shell_id + 2 * thk)
            b.formula(ws, f"W{r}", f"=ROUND({barg}*1.01972,2)", excel_round(barg * 1.01972, 2))
            for col in ("Y", "Z", "AA"):
                ws[f"{col}{r}"] = "No"
            r += 1
        if r - first > 1:
            for col in "ABCDEFGHI":
                ws.merge_cells(f"{col}{first}:{col}{r - 1}")
                ws[f"{col}{first}"].alignment = Alignment(vertical="top", wrap_text=True)
    ws["B2"].comment = Comment("Practice enquiry: two bottles of each size.", "KRONOS template")
    ws["W5"].comment = Comment("Design pressure 12 barg, converted at 1 bar = 1.01972 kg/cm2.", "KRONOS template")
    ws.freeze_panes = "B2"
    return b.save()


# ---------------------------------------------------------------- the Word templates

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DOC_NS = (f'xmlns:w="{W_NS}" xmlns:r="{R_NS}" '
          'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
          'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'


def run(text, bold=False, size=None, color=None):
    props = ("<w:b/>" if bold else "") + (f'<w:color w:val="{color}"/>' if color else "") + \
        (f'<w:sz w:val="{size}"/>' if size else "")
    rpr = f"<w:rPr>{props}</w:rPr>" if props else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def para(*runs, style=None, align=None):
    ppr = ""
    if style or align:
        ppr = "<w:pPr>" + (f'<w:pStyle w:val="{style}"/>' if style else "") + \
            (f'<w:jc w:val="{align}"/>' if align else "") + "</w:pPr>"
    body = "".join(r if r.startswith("<w:") else run(r) for r in runs)
    return f"<w:p>{ppr}{body}</w:p>"


def marker(text):
    """A {{BEGIN name}} or {{END name}} line: its own paragraph, removed when the document is filled."""
    return para(run(text, color="7F7F7F"))


def cell(content, width, head=False):
    shade = '<w:shd w:val="clear" w:color="auto" w:fill="DDEBF7"/>' if head else ""
    inner = content if content.startswith("<w:p") else para(run(content, bold=head))
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{shade}</w:tcPr>{inner}</w:tc>'


def table(rows, widths, head_rows=1):
    border = "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                     for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in widths)
    out = [f'<w:tbl><w:tblPr><w:tblW w:w="{sum(widths)}" w:type="dxa"/><w:tblBorders>{border}</w:tblBorders>'
           f'</w:tblPr><w:tblGrid>{grid}</w:tblGrid>']
    for n, row in enumerate(rows):
        out.append("<w:tr>" + "".join(cell(c, w, n < head_rows) for c, w in zip(row, widths)) + "</w:tr>")
    out.append("</w:tbl>")
    return "".join(out)


def page_field(code):
    return (f'<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> {code} '
            '</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>')


def picture_run(rid, cx, cy, pid):
    return (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
            f'<wp:docPr id="{pid}" name="Logo {pid}"/><a:graphic><a:graphicData '
            'uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr>'
            f'<pic:cNvPr id="{pid}" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill>'
            f'<a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm>'
            f'<a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/>'
            '</a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')


STYLES = (XML_HEAD + f'<w:styles xmlns:w="{W_NS}"><w:docDefaults><w:rPrDefault><w:rPr>'
          '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault>'
          '</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
          '<w:pPr><w:spacing w:after="80"/></w:pPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>'
          '<w:pPr><w:spacing w:before="200" w:after="80"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>'
          '<w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="360"/></w:pPr></w:style></w:styles>')


def docx(title, body, header, footer):
    """A Word file: the body, one header with the logo, one footer; every part written here, so the bytes are fixed."""
    sect = ('<w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/>'
            '<w:footerReference w:type="default" r:id="rIdFooter"/><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" '
            'w:gutter="0"/></w:sectPr>')
    parts = {
        "[Content_Types].xml": (
            XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.'
            'wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
            '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.'
            'wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/'
            'vnd.openxmlformats-package.core-properties+xml"/></Types>'),
        "_rels/.rels": rels_xml([
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
            'officeDocument" Target="word/document.xml"/>',
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/'
            'core-properties" Target="docProps/core.xml"/>']).decode("utf-8"),
        "docProps/core.xml": CORE.format(title=escape(title)),
        "word/document.xml": XML_HEAD + f"<w:document {DOC_NS}><w:body>{body}{sect}</w:body></w:document>",
        "word/styles.xml": STYLES,
        "word/header1.xml": XML_HEAD + f"<w:hdr {DOC_NS}>{header}</w:hdr>",
        "word/footer1.xml": XML_HEAD + f"<w:ftr {DOC_NS}>{footer}</w:ftr>",
        "word/_rels/document.xml.rels": rels_xml([
            '<Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
            'styles" Target="styles.xml"/>',
            '<Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
            'header" Target="header1.xml"/>',
            '<Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
            'footer" Target="footer1.xml"/>']).decode("utf-8"),
        "word/_rels/header1.xml.rels": rels_xml([
            '<Relationship Id="rIdLogo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
            'image" Target="media/logo.png"/>']).decode("utf-8"),
        "word/media/logo.png": logo_png(),
    }
    parts = {k: v.encode("utf-8") if isinstance(v, str) else v for k, v in parts.items()}
    return write_zip(parts, list(parts))


def header_with(*lines):
    logo = para(picture_run("rIdLogo", 609600, 304800, 1))
    return logo + "".join(para(run(text, bold=n == 0)) for n, text in enumerate(lines))


def footer_with(text):
    return para(run(text + "    Page "), page_field("PAGE"), run(" of "), page_field("NUMPAGES"), align="right")


def estimate_template():
    body = [
        para(run("ESTIMATION SHEET - ENQUIRY {{Enquiry}}", bold=True, size=28)),
        para("{{Title}}"),
        table([["Client", "{{Client}}"],
               # one placeholder split over two runs, as Word often stores what a person typed
               ["End user", para(run("{{End_"), run("User}}"))],
               ["Enquiry", "{{Enquiry}}"],
               ["Price basis", "{{Price_Basis}}"]], [2400, 7200], head_rows=0),
        para(run("1. Price Summary"), style="Heading1"),
        table([["Sl.", "Item", "Qty", "Est. wt (kg)", "Unit price (INR)", "Total (INR)"],
               ["{{Sl}}", "{{Item}}", "{{Qty}}", "{{Weight}}", "{{Unit_Price}}", "{{Total}}"],
               ["", "Total (excl. GST)", "{{Total_Qty}}", "", "", "{{Grand_Total}}"]],
              [600, 3600, 700, 1300, 1700, 1700]),
        para(run("2. Basis & Assumptions"), style="Heading1"),
        marker("{{BEGIN Assumption}}"),
        para("{{Text}}", style="ListParagraph"),
        marker("{{END Assumption}}"),
        para(run("3. Weight Build-up per Item (kg)"), style="Heading1"),
        # the second column repeats once per item
        table([["Parameter", "{{Tag}}"],
               ["Shell", "{{Shell_Wt}}"],
               ["Dished ends / flanges", "{{Ends_Wt}}"],
               ["Finished weight per item", "{{Fin_Wt}}"]], [4000, 1800]),
        para(run("4. Cost Build-up per Item"), style="Heading1"),
        marker("{{BEGIN Cost}}"),
        para(run("4.{{Sl}}  Cost Build-up - {{Item}} (per No.)", bold=True)),
        table([["Cost head", "Amount (INR)"],
               ["A. Materials", "{{Materials}}"],
               ["B. Conversion", "{{Conversion}}"],
               ["Factory cost (A + B)", "{{Factory_Cost}}"],
               ["Unit price (rounded up to INR 1,000)", "{{Unit_Price}}"]], [6000, 2400]),
        marker("{{END Cost}}"),
        para(""),   # two tables with nothing between them show as one in Word
        table([["Prepared by", "Checked by", "Approved by"],
               ["Name:", "Name:", "Name:"]], [3200, 3200, 3200]),
        para(""),   # Word ends a body with a paragraph
    ]
    header = header_with("Estimation Sheet No. {{Est_No}}    Date: {{Date}}",
                         "INTERNAL - PRACTICE TEMPLATE - NOT FOR A CLIENT")
    return docx("Internal estimation sheet (look-alike)", "".join(body), header,
                footer_with("Practice internal estimate - invented rates"))


def proposal_template():
    body = [
        para(run("Part A - Technical Proposal", bold=True, size=32), align="center"),
        para("Submitted to", align="center"),
        table([["Client", ":", "{{Client}}"],
               ["Item", ":", "{{Item_Line}}"],
               ["End User / Project", ":", "{{End_User}}"],
               ["Your Enquiry Ref.", ":", "{{Enquiry_Ref}}"],
               ["Proposal No.", ":", "{{Proposal_No}}"],
               ["Date", ":", "{{Date}}"]], [2600, 300, 6700], head_rows=0),
        para("By"),
        para(run("DEMO FABRICATORS PRIVATE LIMITED", bold=True)),
        para(run("Annexure - A    Basis of Offer"), style="Heading1"),
        para("{{Basis}}"),
        table([["SL. No.", "Item", "Tag No.", "Qty"],
               ["{{Sl}}", "{{Item}}", "{{Tag}}", "{{Qty}}"],
               ["", "Total quantity", "", "{{Total_Qty}}"]], [900, 5100, 2000, 1000]),
        para(run("Annexure - B    Technical Delivery Conditions"), style="Heading1"),
        marker("{{BEGIN Equipment}}"),
        para(run("Annexure - B (contd.)  Equipment {{Sl}}", bold=True)),
        table([["Technical Delivery Conditions", ""],
               ["Equipment Name", "{{Item}}"],
               ["Tag Nos / Quantity", "{{Tag}} / {{Qty}} Nos"],
               ["Shell ID x Length (TL to TL) x Thk", "{{Shell_ID}} ID x {{TL_TL}} TL-TL x {{Thk}} thk"],
               ["Design Pressure", "{{Design_Pr}}"]], [4000, 5600]),
        # a repeated row inside the repeated section: one row per nozzle of this item
        table([["Mark", "Description", "Size (NB)", "Rating / Class"],
               ["{{Mark}}", "{{Nozzle}}", "{{Size}}", "{{Rating}}"]], [1200, 4400, 1600, 2400]),
        para("Nozzle notes: {{Nozzle_Notes}}"),
        marker("{{END Equipment}}"),
        marker("{{BEGIN Nace}}"),
        para(run("Annexure - C (F)    NACE / Wet H2S Requirements"), style="Heading1"),
        table([["Requirement", "KGC confirmation"],
               ["{{Requirement}}", "{{Confirmation}}"]], [4800, 4800]),
        marker("{{END Nace}}"),
        para(run("Annexure - E    Technical Queries & Deviations"), style="Heading1"),
        table([["TQ No.", "Inquiry requirement / observation", "Proposal / query"],
               ["{{TQ_No}}", "{{Requirement}}", "{{Query}}"]], [1200, 4200, 4200]),
        para(run("Annexure - G    General Terms & Conditions"), style="Heading1"),
        para(run("1. Effective Date of Contract", bold=True)),
        para("The contract is effective from the date of the purchase order and its acceptance (practice text)."),
    ]
    header = header_with("Technical Proposal No. {{Proposal_No}}    Date: {{Date}}")
    return docx("Technical proposal (look-alike)", "".join(body), header, footer_with("QF-DM-009 Rev 1"))


BUILDERS = {"wps_form.xlsx": wps_form, "wps_form_b.xlsx": wps_form_b, "estimation.xlsx": estimation, "enquiry_register.xlsx": enquiry_register,
            "estimate_template.docx": estimate_template, "proposal_template.docx": proposal_template}


def build(name):
    return BUILDERS[name]()


def main():
    for name in FILES:
        (HERE / name).write_bytes(build(name))
        print(f"wrote data/templates/{name}")


if __name__ == "__main__":
    sys.exit(main())
