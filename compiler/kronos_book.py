"""KRONOS Offline's workbook copy (D87, docs/kronos-offline.md): load an uploaded .xlsx or .xlsm, run a proven recipe's
procedure on it with the reference interpreter, and write the changed copy. The uploaded file is never changed: the
copy is new bytes, and the interpreter's meaning is the one every practice TEST proves.

    describe(data, name)                  # sheets, headings, sample rows, and what the file holds (for the AI's text)
    arguments(text)                       # the recipe's ARGS with a default taken from its first TEST
    apply(text, data, name, args, today)  # (a result dict, the changed workbook's bytes or None)

How a cell survives. The interpreter holds values, fills, number formats and hidden rows; it never holds fonts, borders,
widths or formulas' text. So every loaded cell carries an origin, (sheet, row, col, keep_style), which the interpreter
moves with the cell (delete, insert, sort, COPY_SHEET). Writing the copy rebuilds each changed sheet cell by cell:
the cell's own style from its origin, its own formula while its value is unchanged, the interpreter's value, fill and
number format where they differ from what was loaded. A sheet the recipe didn't change is never touched.

openpyxl is imported when first needed, so the toolchain loads without it.
"""
import datetime as dt
import io
import re
import sys
import warnings
import zipfile
from copy import copy
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "vcl"))
import semantics as S  # noqa: E402
from errcodes import coded  # noqa: E402
from vcl_fixtures import Sheet  # noqa: E402
from vcl_interp import Book, Formula, Machine, VclError, entry_of, params_of, proc_name  # noqa: E402
from vcl_parse import Token, Tree, parse  # noqa: E402

SAMPLE_ROWS = 10        # rows per sheet in the text the AI gets
SAMPLE_TEXT = 60        # characters per cell in that text
MAX_CHANGES = 50        # cells listed in "the first 50 changes"
FILE_OPS = ("MAKE_FOLDER", "COPY_FILE", "LIST_FILES", "READ_TEXT", "WRITE_TEXT", "FILE_EXISTS", "FILE_COUNT",
            "FOLDER_EXISTS", "PATH_JOIN")


class BookError(Exception):
    """Something about the file or the recipe that stops the copy, in words for the person."""


# ---------------------------------------------------------------- what a file holds (zip parts)

PARTS = (   # (plain name, how many of them: a count over the zip's names and parts)
    ("chart", lambda names, read: sum(bool(re.fullmatch(r"xl/charts/chart\d+\.xml", n)) for n in names)),
    ("picture", lambda names, read: sum(n.startswith("xl/media/") for n in names)),
    ("pivot table", lambda names, read: sum(bool(re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", n)) for n in names)),
    ("slicer", lambda names, read: sum(bool(re.fullmatch(r"xl/slicers/slicer\d+\.xml", n)) for n in names)),
    ("threaded comment", lambda names, read: sum(bool(re.fullmatch(r"xl/threadedComments/.*\.xml", n)) for n in names)),
    ("form control", lambda names, read: sum(n.startswith("xl/ctrlProps/") for n in names)),
    ("ActiveX control", lambda names, read: sum(bool(re.fullmatch(r"xl/activeX/activeX\d+\.xml", n)) for n in names)),
    ("embedded object", lambda names, read: sum(n.startswith("xl/embeddings/") for n in names)),
    ("link to another workbook", lambda names, read: sum(bool(re.fullmatch(r"xl/externalLinks/externalLink\d+\.xml", n))
                                                       for n in names)),
    ("sparkline group", lambda names, read: sum(len(re.findall(r"<x14:sparklineGroup[ >]", read(n))) for n in names
                                                if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n))),
    ("shape or text box", lambda names, read: sum(len(re.findall(r"<xdr:(?:sp|grpSp|cxnSp)[ >]", read(n))) for n in names
                                                  if re.fullmatch(r"xl/drawings/drawing\d+\.xml", n))),
)
SAVED_AGAIN = ("chart", "pivot table")   # openpyxl keeps these but writes them its own way


def parts_of(data):
    """{plain name: count} of what the zip holds, for the parts above (only those present)."""
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        names = z.namelist()

        def read(n):
            return z.read(n).decode("utf-8", "replace")
        found = {name: count(names, read) for name, count in PARTS}
    return {k: v for k, v in found.items() if v}


def plural(n, word):
    if n == 1:
        return f"1 {word}"
    return f"{n} {word}es" if word.endswith(("x", "ch")) else f"{n} {word}s"


# ---------------------------------------------------------------- loading

def _openpyxl():
    try:
        import openpyxl
    except ImportError as ex:   # a Python that was started without the wheel
        raise BookError("KRONOS can't read workbooks here: its workbook reader (openpyxl) isn't loaded") from ex
    return openpyxl


def check_file(data, name):
    """Refuse what isn't an .xlsx or .xlsm, in words the person can act on."""
    lower = (name or "").lower()
    if lower.endswith(".xls"):
        raise BookError("This is an old .xls workbook. In Excel choose File, Save As, and save it as .xlsx; then upload that.")
    if data[:4] == b"\xd0\xcf\x11\xe0":
        raise BookError("This file is an old .xls workbook or is protected with a password. Save it in Excel as .xlsx "
                        "without a password, and upload that.")
    if not zipfile.is_zipfile(io.BytesIO(data)):
        raise BookError("This doesn't look like an Excel workbook (.xlsx or .xlsm).")
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        if "xl/workbook.xml" not in z.namelist():
            raise BookError("This is a zip file but not an Excel workbook (.xlsx or .xlsm).")


def _to_model(cell):
    """One loaded value as the interpreter holds it."""
    v, kind = cell.value, cell.data_type
    if v is None:
        return None
    if kind == "e":
        return S.XlError(str(v))
    if isinstance(v, dt.datetime):
        return v.date() if (v.hour, v.minute, v.second, v.microsecond) == (0, 0, 0, 0) else v
    if isinstance(v, dt.time):
        return (v.hour * 3600 + v.minute * 60 + v.second + v.microsecond / 1e6) / 86400
    if isinstance(v, dt.timedelta):
        return v.total_seconds() / 86400
    if isinstance(v, (bool, int, float, str, dt.date)):
        return v
    return str(v)   # rich text and anything else: its text


def _fill_of(cell):
    f = cell.fill
    if f is not None and f.fill_type == "solid" and f.fgColor is not None and f.fgColor.type == "rgb":
        rgb = str(f.fgColor.rgb)
        if len(rgb) == 8 and rgb.upper() != "00000000":
            return "#" + rgb[2:].upper()
    return None


class Loaded:
    """An uploaded workbook: the openpyxl workbook that will be written, the interpreter's view of it, and what the
    copy needs from the original (each cell's style, formula and value; each row's height)."""

    def __init__(self, data, name):
        check_file(data, name)
        openpyxl = _openpyxl()
        self.data, self.name = data, name
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            self.xlsm = (name or "").lower().endswith(".xlsm") or "xl/vbaProject.bin" in z.namelist()
        try:
            with warnings.catch_warnings(record=True) as said:   # openpyxl says what it will drop
                warnings.simplefilter("always")
                self.wb = openpyxl.load_workbook(io.BytesIO(data), keep_vba=self.xlsm)
                values = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        except zipfile.BadZipFile as ex:
            raise BookError("This doesn't look like an Excel workbook (.xlsx or .xlsm).") from ex
        except Exception as ex:   # openpyxl names many ways a file can be broken
            raise BookError(f"KRONOS couldn't read this workbook ({type(ex).__name__}: {ex}).") from ex
        self.no_result = 0   # formula cells the file holds no saved result for
        self.formula_cells = 0
        self.sheets = {}     # name -> Sheet (the interpreter's)
        self.initial = {}    # name -> Sheet snapshot as loaded
        self.cells = {}      # name -> {(row, col): (style, formula, value, model value)}
        self.rowdims = {}    # name -> {row: (height, outline level)}
        for ws in self.wb.worksheets:
            sh = Sheet(ws.title)
            sh.origins = {}
            kept = {}
            for (r, c), cell in ws._cells.items():
                if type(cell).__name__ == "MergedCell":
                    continue
                cached = values[ws.title]._cells.get((r, c))
                is_formula = cell.data_type == "f"
                model = _to_model(cached) if cached is not None else (None if is_formula else _to_model(cell))
                if is_formula:
                    self.formula_cells += 1
                    if model is None:
                        self.no_result += 1
                if model is not None:
                    sh.cells[(r, c)] = model
                fill = _fill_of(cell)
                if fill:
                    sh.fills[(r, c)] = fill
                if cell.number_format and cell.number_format != "General":
                    sh.formats[(r, c)] = cell.number_format
                sh.origins[(r, c)] = (ws.title, r, c, True)
                extras = (copy(cell.comment) if cell.comment else None, copy(cell.hyperlink) if cell.hyperlink else None)
                kept[(r, c)] = (copy(cell._style), cell.value if is_formula else None, cell.value, model, extras)
            heights = {}
            for r, dim in ws.row_dimensions.items():
                if dim.hidden:
                    sh.hidden.add(r)
                if dim.height is not None or dim.outlineLevel:
                    heights[r] = (dim.height, dim.outlineLevel)
            self.sheets[ws.title] = sh
            self.initial[ws.title] = sh.snapshot()
            self.cells[ws.title] = kept
            self.rowdims[ws.title] = heights
        self.parts = parts_of(data)
        dropped = []   # "Data Validation extension is not supported and will be removed" and its kind
        for w in said:
            m = re.fullmatch(r"(.+?) extension is not supported and will be removed", str(w.message))
            if m and "sparkline" not in m.group(1).lower() and m.group(1) not in dropped:
                dropped.append(m.group(1))
        self.dropped = dropped

    def book(self):
        return Book({name: sh.snapshot() for name, sh in self.sheets.items()})


# ---------------------------------------------------------------- describe

def _show(v):
    text = S.text_of(v)
    return text if len(text) <= SAMPLE_TEXT else text[:SAMPLE_TEXT - 1] + "…"


def describe(data, name):
    """The sheets of an uploaded workbook, for the page and for the text the person copies to the AI chat."""
    book = Loaded(data, name)
    sheets = []
    for sh in book.sheets.values():
        rows, cols = sh.max_row() if sh.cells else 0, sh.max_col() if sh.cells else 0
        headings = [_show(sh.get(1, c)) for c in range(1, cols + 1)]
        sample = [[_show(sh.get(r, c)) for c in range(1, cols + 1)] for r in range(2, min(rows, SAMPLE_ROWS + 1) + 1)]
        sheets.append({"name": sh.name, "rows": rows, "cols": cols, "headings": headings, "sample": sample})
    features = [f"{plural(n, word)}" for word, n in book.parts.items()]
    notes = []
    if book.no_result:
        notes.append(f"{plural(book.no_result, 'formula cell')} {'has' if book.no_result == 1 else 'have'} no saved result "
                     "(the file was never calculated in Excel), so KRONOS sees them as empty: open and save the file "
                     "in Excel first if the recipe reads them.")
    lines = [f"The workbook {name} has {plural(len(sheets), 'sheet')}:", ""]
    for s in sheets:
        lines.append(f"Sheet \"{s['name']}\": {s['rows']} rows (row 1 is the headings), {s['cols']} columns")
        lines.append("  Headings: " + " | ".join(f"{i + 1}={h or '(blank)'}" for i, h in enumerate(s["headings"])))
        if s["sample"]:
            shown = len(s["sample"])
            lines.append(f"  The first {shown} data row{'' if shown == 1 else 's'}:" if s["rows"] - 1 > shown
                         else f"  The {shown} data row{'' if shown == 1 else 's'}:")
            lines += ["    " + " | ".join(row) for row in s["sample"]]
        lines.append("")
    return {"name": name, "sheets": sheets, "features": features, "xlsm": book.xlsm, "notes": notes,
            "text": "\n".join(lines).rstrip() + "\n"}


# ---------------------------------------------------------------- the recipe's arguments

def test_call(k):
    """What a TEST passes and how many times it runs it: ([the call's values], runs). Names and numbers as they are,
    text without its quotes."""
    values, runs = [], 1
    rest = k["tree"].children[2:]
    for i, item in enumerate(rest):
        if isinstance(item, Token) and item.type == "INT":
            runs = int(item)
            break
        if isinstance(item, Tree) and item.data == "assertion":
            break
        values.append(str(item)[1:-1] if isinstance(item, Token) and item.type == "STRING" else str(item))
    return values, runs


def _call_values(keys):
    """The values the first TEST passes."""
    for k in keys:
        if k["key"] == "TEST":
            return test_call(k)[0]
    return []


def arguments(text):
    """[{name, type, default}] of the recipe's procedure (a BUILD's modMain.RunHeadless)."""
    spec, problems = parse(text)
    if problems or not spec:
        raise BookError("The recipe doesn't parse.")
    entry, _, keys = entry_of(spec)
    values = _call_values(keys)
    return [{"name": n, "type": t, "default": values[i] if i < len(values) else ""}
            for i, (n, t) in enumerate(params_of(entry))]


def uses_files(text):
    """True when the recipe works with files or folders (an op, a folder FIXTURE or MODIFIES ... files): that needs a
    real folder, which a workbook copy doesn't have."""
    for line in text.splitlines():
        key, _, rest = line.strip().partition(" ")
        if key == "OP" and re.search(r"\b(" + "|".join(FILE_OPS) + r")\b", rest):
            return True
        if key == "FIXTURE" and re.search(r"\bfolder=", rest):
            return True
        if key == "MODIFIES" and re.search(r"\bfiles\b", rest):
            return True
    return False


def _bind(entry, given, book):
    """The ARGS the person chose, as the values the interpreter takes."""
    bound = {}
    for pname, typ in params_of(entry):
        text = str(given.get(pname, "")).strip() if given.get(pname) is not None else ""
        if typ == "Workbook":
            bound[pname] = book
        elif typ == "Worksheet":
            if not text:
                raise BookError(f"Choose a sheet for {pname}.")
            found = next((sh for key, sh in book.sheets.items() if key.casefold() == text.casefold()), None)
            if found is None:
                raise BookError(f"This workbook has no sheet named {text!r}.")
            bound[pname] = found
        elif typ == "String":
            bound[pname] = str(given.get(pname, ""))
        elif typ in ("Long", "Double"):
            try:
                bound[pname] = S.to_number(text)
            except Exception as ex:
                raise BookError(f"{pname} needs a number, not {text!r}.") from ex
            if typ == "Long" and not float(bound[pname]).is_integer():
                raise BookError(f"{pname} needs a whole number, not {text!r}.")
        elif typ == "Boolean":
            if text.upper() not in ("TRUE", "FALSE"):
                raise BookError(f"{pname} needs TRUE or FALSE.")
            bound[pname] = text.upper() == "TRUE"
        elif typ == "Date":
            try:
                bound[pname] = dt.date.fromisoformat(text)
            except ValueError as ex:
                raise BookError(f"{pname} needs a date written like 2026-11-14.") from ex
        else:
            raise BookError(f"This recipe takes {pname} As {typ}, which KRONOS can't ask for here.")
    return bound


# ---------------------------------------------------------------- writing the copy

R1C1 = re.compile(r"(?<![A-Za-z0-9_.])R(\[-?\d+\]|\d+)?C(\[-?\d+\]|\d+)?(?![A-Za-z0-9_(\[])")
R1C1_LEFT = re.compile(r"(?<![A-Za-z0-9_.])[RC](\[-?\d+\]|\d+)?(?![A-Za-z0-9_(\[])")
QUOTED = re.compile(r'"(?:[^"]|"")*"|\'(?:[^\']|\'\')*\'')


def _letters(c):
    out = ""
    while c:
        c, rem = divmod(c - 1, 26)
        out = chr(65 + rem) + out
    return out


def r1c1_to_a1(text, row, col):
    """A SET_FORMULA style=r1c1 formula as A1 text for the cell it goes in. Only cell references are understood."""
    def one(part, here, absolute_prefix):
        if part is None:
            return str(here), False
        if part.startswith("["):
            return str(here + int(part[1:-1])), False
        return part, True

    def cell(m):
        r_text, r_abs = one(m.group(1), row, "$")
        c_text, c_abs = one(m.group(2), col, "$")
        r_no, c_no = int(r_text), int(c_text)
        if not (1 <= r_no <= 1048576 and 1 <= c_no <= 16384):
            raise BookError(f"The formula {text} points outside the sheet from row {row}, column {col}.")
        return f"{'$' if c_abs else ''}{_letters(c_no)}{'$' if r_abs else ''}{r_no}"

    pieces, last = [], 0
    for m in QUOTED.finditer(text):   # leave text and sheet names alone
        pieces.append((False, text[last:m.start()]))
        pieces.append((True, m.group()))
        last = m.end()
    pieces.append((False, text[last:]))
    out = []
    for quoted, piece in pieces:
        if quoted:
            out.append(piece)
            continue
        if R1C1_LEFT.search(R1C1.sub("@", piece)):   # a row or column on its own, which has no A1 cell
            raise BookError(f"KRONOS can't copy the formula {text}: it uses a whole row or column reference. Use the macro.")
        out.append(R1C1.sub(cell, piece))
    return "".join(out)


def _same(a, b):
    if isinstance(a, bool) != isinstance(b, bool):
        return False
    return a == b


def _sheet_changed(sh, initial):
    return not (sh.cells == initial.cells and sh.fills == initial.fills and sh.formats == initial.formats
                and sh.hidden == initial.hidden)


def _write_sheet(loaded, sh, ws):
    """Rebuild one worksheet from the interpreter's sheet."""
    from openpyxl.styles import PatternFill, is_date_format
    merged = [str(rng) for rng in ws.merged_cells.ranges]
    for rng in merged:
        ws.unmerge_cells(rng)
    ws._cells.clear()
    ws.row_dimensions.clear()
    origins = sh.origins or {}
    positions = set(sh.cells) | set(sh.fills) | set(sh.formats) | {rc for rc, o in origins.items() if o[3]}
    row_origin = {}   # final row -> (sheet, row) its first cell came from
    for (r, c) in sorted(origins):
        row_origin.setdefault(r, origins[(r, c)][:2])
    for (r, c) in sorted(positions):
        value = sh.cells.get((r, c))
        o = origins.get((r, c))
        kept = loaded.cells[o[0]].get(o[1:3]) if o else None
        style = kept[0] if kept and o[3] else None
        cell = ws.cell(row=r, column=c)
        if isinstance(value, Formula):
            cell.value = value.text if not value.r1c1 else r1c1_to_a1(value.text, r, c)
            if not cell.value.startswith("="):
                cell.value = "=" + cell.value
        elif kept and kept[1] is not None and _same(value, kept[3]) and o[3]:
            cell.value = kept[1]
        elif value is not None:
            if isinstance(value, S.XlError):
                cell.value = value.text
                cell.data_type = "e"
            else:
                cell.value = value
                if isinstance(value, str):
                    cell.data_type = "s"
        if style is not None:
            cell._style = copy(style)
        if kept and not (value is None and not o[3]):   # a note and a link stay with their cell (CLEAR all removes them)
            comment, link = kept[4]
            if comment is not None:
                cell.comment = copy(comment)
            if link is not None:
                cell.hyperlink = copy(link)
        initial = loaded.initial[o[0]] if o else None
        was_fill = initial.fills.get(o[1:3]) if o and o[3] else None
        was_fmt = initial.formats.get(o[1:3], "General") if o and o[3] else "General"
        fill, fmt = sh.fills.get((r, c)), sh.formats.get((r, c), "General")
        if fill != was_fill:
            cell.fill = PatternFill(fill_type="solid", start_color="FF" + fill[1:], end_color="FF" + fill[1:]) if fill \
                else PatternFill(fill_type=None)
        if fmt != was_fmt:
            cell.number_format = fmt
        if isinstance(value, (dt.date, dt.datetime)) and not is_date_format(cell.number_format):
            cell.number_format = "yyyy-mm-dd" if not isinstance(value, dt.datetime) else "yyyy-mm-dd hh:mm:ss"
    rows = set(row_origin) | sh.hidden | {r for r, _ in positions}
    for r in rows:   # a row keeps the height of the row its first cell came from, and FILTER's hidden state
        dim = ws.row_dimensions[r]
        if r in row_origin:
            height, level = loaded.rowdims.get(row_origin[r][0], {}).get(row_origin[r][1], (None, 0))
            if height is not None:
                dim.height = height
            if level:
                dim.outlineLevel = level
        dim.hidden = r in sh.hidden
    for rng in merged:
        ws.merge_cells(rng)


def _order(loaded, book_model):
    """Put the worksheets in the interpreter's order; a chart sheet keeps its place among the slots."""
    wb = loaded.wb
    sheets = {ws.title.casefold(): ws for ws in wb._sheets if hasattr(ws, "cell")}
    wanted = [sheets[name.casefold()] for name in book_model.sheets if name.casefold() in sheets]
    slots = [i for i, s in enumerate(wb._sheets) if hasattr(s, "cell")]
    for i, ws in zip(slots, wanted):
        wb._sheets[i] = ws
    for ws in wanted[len(slots):]:
        wb._sheets.append(ws)


def _write(loaded, model):
    """The changed workbook's bytes, and which sheets changed."""
    _openpyxl()
    wb = loaded.wb
    status = {}
    existing = {ws.title: ws for ws in wb.worksheets}
    for name, sh in model.sheets.items():
        ws = existing.get(name)
        if ws is None:
            if sh.copy_of and sh.copy_of in existing:
                ws = wb.copy_worksheet(existing[sh.copy_of])
                ws.title = name
            else:
                ws = wb.create_sheet(name)
            existing[name] = ws
            status[name] = "new"
            _write_sheet(loaded, sh, ws)
        elif _sheet_changed(sh, loaded.initial[name]):
            status[name] = "changed"
            _write_sheet(loaded, sh, ws)
        else:
            status[name] = "unchanged"
    _order(loaded, model)
    out = io.BytesIO()
    try:
        wb.save(out)
    finally:
        archive = getattr(wb, "vba_archive", None)   # openpyxl leaves a macro workbook's zip open
        if archive is not None:
            try:
                archive.close()
            except ValueError:
                pass
    return out.getvalue(), status


# ---------------------------------------------------------------- what changed, and what to warn about

def _row_story(sh, initial):
    """(rows deleted, rows added, rows moved) of an uploaded sheet, from where its cells came from."""
    origins = sh.origins or {}
    kept = {o[1] for o in origins.values() if o[0] == sh.name}
    deleted = sorted({r for r, _ in initial.cells} - kept)
    with_origin = {r for r, _ in origins}
    added = sorted({r for r, _ in sh.cells} - with_origin)
    first = {}
    for (r, c) in sorted(origins):
        if sh.cells.get((r, c)) is not None:
            first.setdefault(r, origins[(r, c)])
    moved = sum(1 for r, o in first.items() if o[0] == sh.name and o[1] != r)
    return deleted, added, moved


def _summary(model, status, loaded):
    sheets, changes = [], []
    total = 0
    for name, sh in model.sheets.items():
        entry = {"name": name, "status": status[name], "cells": 0, "rows_deleted": 0, "rows_added": 0, "rows_moved": 0,
                 "fills": 0, "hidden": len(sh.hidden)}
        if status[name] == "new":
            entry["cells"] = len(sh.cells)
        elif status[name] == "changed":
            initial = loaded.initial[name]
            keys = sorted(set(sh.cells) | set(initial.cells))
            for rc in keys:
                a, b = initial.cells.get(rc), sh.cells.get(rc)
                if not (_same(a, b) or (isinstance(a, Formula) and a == b)):
                    entry["cells"] += 1
                    total += 1
                    if len(changes) < MAX_CHANGES:
                        changes.append([name, f"{_letters(rc[1])}{rc[0]}", S.text_of(a) if not isinstance(a, Formula)
                                        else a.text, S.text_of(b) if not isinstance(b, Formula) else b.text])
            entry["fills"] = sum(sh.fills.get(rc) != initial.fills.get(rc) for rc in set(sh.fills) | set(initial.fills))
            deleted, added, moved = _row_story(sh, initial)
            entry.update(rows_deleted=len(deleted), rows_added=len(added), rows_moved=moved)
        sheets.append(entry)
    return sheets, changes, total


def _warnings(loaded, model, status, out_bytes):
    found = []
    after = parts_of(out_bytes)
    for word, n in loaded.parts.items():
        left = after.get(word, 0)
        if left < n:
            lost = n - left
            found.append({"kind": "lost", "text": f"{plural(lost, word)} in the original will not be in the copy."})
        elif word in SAVED_AGAIN:
            found.append({"kind": "saved again", "text": f"{plural(n, word)} {'is' if n == 1 else 'are'} saved again by "
                             "KRONOS and may look a little different (colours, fonts, spacing)."})
    for kind in loaded.dropped:
        found.append({"kind": "lost", "text": f"{kind} features of the original (Excel 2010 and later) will not be in the copy."})
    for name, sh in model.sheets.items():
        if status[name] != "changed" or name not in loaded.initial:
            continue
        deleted, added, moved = _row_story(sh, loaded.initial[name])
        if not (deleted or added or moved):
            continue
        ws = loaded.wb[name]
        stale = []
        if loaded.formula_cells:
            stale.append("formulas (on this sheet or pointing at it)")
        if ws.merged_cells.ranges:
            stale.append("merged cells")
        if ws.conditional_formatting:
            stale.append("conditional formatting")
        if ws.data_validations and ws.data_validations.dataValidation:
            stale.append("data validation")
        if ws.tables:
            stale.append("tables")
        if ws.auto_filter and ws.auto_filter.ref:
            stale.append("the filter")
        if stale:
            found.append({"kind": "moved", "text": f"Rows were deleted, added or moved on the sheet {name}. Excel would have "
                             f"adjusted {', '.join(stale)}; KRONOS doesn't, so check them in the copy."})
    if loaded.no_result:
        found.append({"kind": "load", "text": f"{plural(loaded.no_result, 'formula cell')} had no saved result, so the recipe saw "
                         "them as empty. Open and save the file in Excel first if the recipe reads them."})
    return found


def _download_name(name):
    stem, dot, ext = (name or "workbook.xlsx").rpartition(".")
    if not dot:
        stem, ext = name, "xlsx"
    return stem, ext.lower() if ext.lower() in ("xlsx", "xlsm") else "xlsx"


# ---------------------------------------------------------------- apply

def apply(text, data, name, args, today=None):
    """Run the proven recipe's procedure on the uploaded workbook and write the changed copy.
    Returns (result, bytes): bytes is None when nothing is given back, and the result says why."""
    from kronos_engine import prove   # here, so that importing this module stays light
    if isinstance(today, str):
        today = dt.date.fromisoformat(today)
    proof = prove(text, today)
    if not proof.ok:
        return _refused("This recipe isn't proven, so KRONOS won't run it on your workbook."), None
    return apply_proven(proof.spec, data, name, args, today)


def _refused(words, **more):
    return {"ok": False, "message": words, "sheets": [], "changes": [], "changes_total": 0, "warnings": [],
            "result": None, "messages": [], **more}


def apply_proven(text, data, name, args, today=None):
    """apply() for a recipe that has already been proven (the page proves it first; the tests prove every spec once)."""
    if isinstance(today, str):
        today = dt.date.fromisoformat(today)
    try:
        spec, problems = parse(text)
        if problems or not spec:
            return _refused("The recipe doesn't parse."), None
        entry, procs, _ = entry_of(spec)
        if uses_files(text):
            return _refused("This recipe works with files or folders, which KRONOS can't do on an uploaded workbook. "
                            "Use the macro instead."), None
        loaded = Loaded(bytes(data), name)
        model = loaded.book()
        bound = _bind(entry, args or {}, model)
        log = []
        machine = Machine(entry, bound, log, today or dt.date.today(), None, model, procs)
        result = machine.run()
    except BookError as ex:
        return _refused(str(ex)), None
    except VclError as ex:
        return _refused(f"The recipe stopped on this workbook: {ex}"), None
    if machine.pre_failed or result == -1:
        return _refused("This workbook doesn't match what the recipe expects, so nothing was changed. "
                        + "; ".join(str(m) for _, _, m in log[:3]), pre_failed=True), None
    if log:
        return _refused("The recipe stopped with an error on this workbook, so nothing was changed: "
                        + "; ".join(str(m) for _, _, m in log[:3]), log=[str(m) for _, _, m in log]), None
    try:
        out, status = _write(loaded, model)
    except BookError as ex:
        return _refused(str(ex)), None
    sheets, changes, total = _summary(model, status, loaded)
    stem, ext = _download_name(name)
    return {"ok": True, "message": "The changes are made on a copy; your original file is untouched.",
            "sheets": sheets, "changes": changes, "changes_total": total,
            "warnings": _warnings(loaded, model, status, out),
            "result": None if result is None else S.text_of(result), "messages": list(machine.messages),
            "stem": stem, "ext": ext}, out
