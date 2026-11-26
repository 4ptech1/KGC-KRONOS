"""KRONOS Offline's workbook copy (D87): compiler/kronos_book.py.

The big check: every proven recipe's TESTs, run on a real .xlsx made from the TEST's own fixture through the upload
path, give the workbook the interpreter itself ends with (values, fills, number formats, hidden rows and sheets). The
other tests cover what the interpreter can't see: fonts, borders, widths, formulas, merged cells, charts, macros.
"""
import datetime as dt
import io
import re
import sys
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(ROOT / "vcl"))

try:
    import openpyxl
    from openpyxl.chart import BarChart, Reference
    from openpyxl.styles import Border, Font, PatternFill, Side, is_date_format
except ImportError:   # CI installs requirements.txt; a bare checkout skips these tests
    openpyxl = None

import semantics as S
from vcl_fixtures import Sheet, build_workbooks
from vcl_interp import Formula, VclError, entry_of, run_tests
from vcl_parse import parse

if openpyxl:
    import kronos_book as kb

TODAY = dt.date(2026, 11, 14)
EXAMPLE = (ROOT / "vcl/examples/example1_rev3.vcl").read_text(encoding="utf-8")


def to_xlsx(sheets):
    """An interpreter workbook ({name: Sheet}) as .xlsx bytes."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, sh in sheets.items():
        ws = wb.create_sheet(name)
        for (r, c), v in sh.cells.items():
            ws.cell(row=r, column=c).value = v.text if isinstance(v, Formula) else (v.text if isinstance(v, S.XlError) else v)
            if isinstance(v, S.XlError):
                ws.cell(row=r, column=c).data_type = "e"
        for (r, c), color in sh.fills.items():
            ws.cell(row=r, column=c).fill = PatternFill(fill_type="solid", start_color="FF" + color[1:], end_color="FF" + color[1:])
        for (r, c), fmt in sh.formats.items():
            ws.cell(row=r, column=c).number_format = fmt
        for r in sh.hidden:
            ws.row_dimensions[r].hidden = True
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


def to_model(cell):
    v = cell.value
    if cell.data_type == "e":
        return S.XlError(str(v))
    if isinstance(v, dt.datetime) and (v.hour, v.minute, v.second) == (0, 0, 0):
        return v.date()
    return v


def read_back(data):
    """{sheet: (cell texts, fills, formats, hidden rows)} of a workbook, as run_tests reports a TEST's state."""
    wb = openpyxl.load_workbook(io.BytesIO(data))
    state = {}
    for ws in wb.worksheets:
        cells, fills, formats = [], [], []
        for (r, c), cell in ws._cells.items():
            v = to_model(cell)
            if v is not None:
                cells.append(((r, c), cell.value if cell.data_type == "f" else S.text_of(v)))
            f = cell.fill
            if f is not None and f.fill_type == "solid":
                fills.append(((r, c), "#" + str(f.fgColor.rgb)[2:].upper()))
            if cell.number_format != "General":
                formats.append(((r, c), cell.number_format, isinstance(v, dt.date)))
        hidden = sorted(r for r, d in ws.row_dimensions.items() if d.hidden)
        state[ws.title] = (sorted(cells), sorted(fills), formats, hidden)
    return state


def digits(cells):
    """Cell texts with numbers cut to the 15 digits Excel keeps: openpyxl writes 16, so the 17th can differ."""
    out = []
    for rc, text in cells:
        try:
            text = format(float(text), ".15g") if re.fullmatch(r"-?\d+\.\d+(e[+-]?\d+)?", text) else text
        except ValueError:
            pass
        out.append((rc, text))
    return out


def call_args(spec_entry, values):
    return {name: v for (name, _), v in zip(kb.params_of(spec_entry), values)}


def sweep_specs():
    paths = sorted((ROOT / "specs").glob("*.vcl")) + sorted((ROOT / "vcl/examples").glob("*.vcl")) \
        + sorted((ROOT / "data/verified").glob("*/spec.vcl"))
    return paths


@unittest.skipIf(openpyxl is None, "openpyxl is not installed")
class RoundTrip(unittest.TestCase):
    """Every TEST of every proven recipe, through an uploaded workbook, equals the interpreter's own final state."""

    def test_every_recipe(self):
        checked = skipped = 0
        for path in sweep_specs():
            text = path.read_text(encoding="utf-8")
            if kb.uses_files(text):
                continue
            spec, problems = parse(text)
            try:
                entry, procs, keys = entry_of(spec)
                results = run_tests(text, TODAY)
            except VclError:   # an older BUILD, whose TESTs the interpreter doesn't run
                continue
            trees = [k["tree"] for k in keys if k["key"] == "FIXTURE"]
            books = build_workbooks(trees, TODAY)
            tests = [k for k in keys if k["key"] == "TEST"]
            self.assertEqual(len(tests), len(results), path)
            for k, expected in zip(tests, results):
                values, runs = kb.test_call(k)
                fixture = str(k["tree"].children[0])
                if fixture not in books:
                    skipped += 1
                    continue
                args = call_args(entry, values)
                sheets = {n: sh.snapshot() for n, sh in books[fixture].items()}
                with self.subTest(spec=path.parent.name if path.name == "spec.vcl" else path.name, line=k["line"]):
                    out = to_xlsx(sheets)
                    read = None
                    if k.get("read"):   # D114, D120: the workbook the TEST reads is the second upload
                        read = (to_xlsx({n: sh.snapshot() for n, sh in books[k["read"]].items()}), k["read"])
                        args = {n: v for n, v in args.items() if v != k["read"]}
                    for _ in range(runs):   # runs=2 runs it again on the copy, as the TEST runs it again on the sheets
                        result, out = kb.apply_proven(text, out, "fixture.xlsx", args, TODAY, read=read)
                        if out is None:
                            break
                    if expected.errors or expected.result == -1:   # a TEST that expects a stop gives no copy
                        self.assertIsNone(out, result["message"])
                        continue
                    self.assertTrue(result["ok"], (result["message"], path))
                    got = read_back(out)
                    self.assertEqual(list(got), list(expected.state), "sheets and their order")
                    for name, (cells, fills, formats) in expected.state.items():
                        g_cells, g_fills, g_formats, g_hidden = got[name]
                        r1c1 = "style=r1c1" in text
                        if r1c1:   # the interpreter keeps the R1C1 text; the copy holds the A1 formula for the cell
                            cells = [(rc, "=" if t.startswith("=") else t) for rc, t in cells]
                            g_cells = [(rc, "=" if t.startswith("=") else t) for rc, t in g_cells]
                        self.assertEqual(digits(g_cells), digits(cells), f"{name}: values")
                        self.assertEqual(g_fills, fills, f"{name}: fills")
                        want = {rc: f for rc, f in formats if f != "General"}
                        have = {rc: f for rc, f, was_date in g_formats if not (was_date and f == "yyyy-mm-dd" and rc not in want)}
                        self.assertEqual(have, want, f"{name}: number formats")
                        self.assertEqual(g_hidden, expected.hidden.get(name, []), f"{name}: hidden rows")
                    checked += 1
        self.assertGreaterEqual(checked, 262, "the sweep must cover every TEST of the repo's recipes (262 today)")
        self.assertEqual(skipped, 0)


def orders(rows=20):
    """A styled orders workbook the way a person might keep one."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Orders"
    ws.append(["Order No", "Order Date", "Customer", "Status"])
    statuses = ["Open", "Cancelled", "Closed", " cancelled "]
    for i in range(1, rows + 1):
        ws.append([i, dt.date(2026, 1, 1) + dt.timedelta(days=i * 3), f"Cust {i}", statuses[i % 4]])
    ws["C5"].font = Font(bold=True, color="FF0000")
    ws["C5"].border = Border(left=Side(style="thin"))
    ws.row_dimensions[5].height = 41
    ws.column_dimensions["C"].width = 33
    ws["A1"].fill = PatternFill(fill_type="solid", start_color="FF00FF00", end_color="FF00FF00")
    return wb


def save(wb):
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


@unittest.skipIf(openpyxl is None, "openpyxl is not installed")
class Copy(unittest.TestCase):
    def apply(self, wb_or_bytes, name="orders.xlsx", text=EXAMPLE, args=None):
        data = wb_or_bytes if isinstance(wb_or_bytes, bytes) else save(wb_or_bytes)
        return kb.apply(text, data, name, args or {"ws": "Orders"}, TODAY), data

    def test_the_example_deletes_and_sorts_and_keeps_styles(self):
        (result, out), original = self.apply(orders())
        self.assertTrue(result["ok"], result["message"])
        self.assertEqual(result["result"], "10")
        sheet = result["sheets"][0]
        self.assertEqual((sheet["status"], sheet["rows_deleted"]), ("changed", 10))
        ws = openpyxl.load_workbook(io.BytesIO(out))["Orders"]
        self.assertEqual(ws.max_row, 11)
        dates = [ws.cell(r, 2).value for r in range(2, 12)]
        self.assertEqual(dates, sorted(dates, reverse=True))
        self.assertFalse(any(str(ws.cell(r, 4).value).strip().lower() == "cancelled" for r in range(2, 12)))
        self.assertTrue(is_date_format(ws["B2"].number_format))
        self.assertEqual(ws.column_dimensions["C"].width, 33)
        row = next(r for r in range(2, 12) if ws.cell(r, 3).value == "Cust 4")   # the styled cell moved with its row
        self.assertTrue(ws.cell(row, 3).font.bold)
        self.assertEqual(ws.cell(row, 3).font.color.rgb, "00FF0000" if ws.cell(row, 3).font.color.rgb.startswith("00") else "FFFF0000")
        self.assertEqual(ws.cell(row, 3).border.left.style, "thin")
        self.assertEqual(ws.row_dimensions[row].height, 41)
        self.assertEqual(ws["A1"].fill.start_color.rgb[2:], "00FF00")

    def test_the_original_is_never_changed(self):
        data = save(orders())
        before = bytes(data)
        result, out = kb.apply(EXAMPLE, data, "orders.xlsx", {"ws": "Orders"}, TODAY)
        self.assertEqual(data, before)
        self.assertNotEqual(out, before)

    def test_notes_links_and_grouping_move_with_their_rows(self):
        from openpyxl.comments import Comment
        wb = orders()
        ws = wb["Orders"]
        ws["C4"].comment = Comment("check this one", "Boss")      # Cust 3 (cancelled: deleted)
        ws["C5"].comment = Comment("a note on Cust 4", "Boss")    # Cust 4 (open: kept)
        ws["C6"].hyperlink = "https://example.com/cust5"           # Cust 5 (cancelled: deleted)
        ws["C7"].hyperlink = "https://example.com/cust6"           # Cust 6 (closed: kept)
        ws.row_dimensions[7].outlineLevel = 1
        (result, out), _ = self.apply(wb)
        ws2 = openpyxl.load_workbook(io.BytesIO(out))["Orders"]
        notes = {c.value: c.comment.text for row in ws2.iter_rows() for c in row if c.comment}
        self.assertEqual(notes, {"Cust 4": "a note on Cust 4"})
        links = {c.value: c.hyperlink.target for row in ws2.iter_rows() for c in row if c.hyperlink}
        self.assertEqual(links, {"Cust 6": "https://example.com/cust6"})
        row = next(r for r in range(2, 12) if ws2.cell(r, 3).value == "Cust 6")
        self.assertEqual(ws2.row_dimensions[row].outlineLevel, 1)

    def test_a_sheet_the_recipe_leaves_alone_is_not_touched(self):
        wb = orders()
        other = wb.create_sheet("Notes")
        other["A1"] = "keep me"
        other["A1"].font = Font(italic=True)
        other.merge_cells("A3:B3")
        (result, out), _ = self.apply(wb)
        self.assertEqual({s["name"]: s["status"] for s in result["sheets"]}, {"Orders": "changed", "Notes": "unchanged"})
        notes = openpyxl.load_workbook(io.BytesIO(out))["Notes"]
        self.assertEqual(notes["A1"].value, "keep me")
        self.assertTrue(notes["A1"].font.italic)
        self.assertEqual([str(r) for r in notes.merged_cells.ranges], ["A3:B3"])

    def test_a_recipe_that_changes_nothing_says_so(self):
        wb = orders()
        ws = wb["Orders"]
        for r in range(ws.max_row, 1, -1):
            if str(ws.cell(r, 4).value).strip().lower() == "cancelled":
                ws.delete_rows(r)
        data = save(wb)
        (result, out), _ = self.apply(data)
        # the sort still runs (newest first), so the sheet changes; no row is deleted
        self.assertEqual(result["sheets"][0]["rows_deleted"], 0)
        self.assertEqual(result["result"], "0")

    def test_formulas_stay_while_their_cell_is_unchanged(self):
        wb = orders()
        ws = wb["Orders"]
        ws["F1"] = "Double"
        for r in range(2, ws.max_row + 1):
            ws.cell(r, 6).value = f"=A{r}*2"
        data = with_cached_values(save(wb), "xl/worksheets/sheet1.xml", {f"F{r}": r - 1 for r in range(2, 22)} | {"F1": None})
        self.assertEqual(kb.Loaded(data, "o.xlsx").no_result, 0)
        text = EXAMPLE.replace("SORT(USED_RANGE($ws), key=cDate, order=desc, header=yes)", "LET(removed, removed)")
        result, out = kb.apply_proven(text, data, "o.xlsx", {"ws": "Orders"}, TODAY)   # deletes only: no row moves
        self.assertTrue(result["ok"], result["message"])
        self.assertEqual(result["sheets"][0]["rows_deleted"], 10)
        ws2 = openpyxl.load_workbook(io.BytesIO(out))["Orders"]
        formulas = [ws2.cell(r, 6).value for r in range(2, ws2.max_row + 1)]
        self.assertTrue(all(isinstance(f, str) and f.startswith("=A") for f in formulas), formulas)
        self.assertTrue(any(w["kind"] == "moved" for w in result["warnings"]))

    def test_a_formula_with_no_saved_result_is_warned_about(self):
        wb = orders()
        wb["Orders"]["F2"] = "=A2*2"
        (result, out), _ = self.apply(wb)
        self.assertTrue(any(w["kind"] == "load" for w in result["warnings"]))

    def test_a_row_op_on_a_sheet_with_merged_cells_is_refused_as_the_macro_refuses_it(self):
        # D113, D119: an uploaded sheet's merged ranges reach the interpreter, so deleting or sorting rows there stops,
        # as ForgeNoMerges stops the macro (before D119 the copy was made with a warning to check them)
        wb = orders()
        wb["Orders"].merge_cells("H1:I1")
        (result, out), _ = self.apply(wb)
        self.assertIsNone(out)
        self.assertIn("merged cells", result["message"])

    def test_a_chart_is_kept_and_said_to_be_saved_again(self):
        wb = orders()
        chart = BarChart()
        chart.add_data(Reference(wb["Orders"], min_col=1, min_row=1, max_row=5))
        wb["Orders"].add_chart(chart, "H2")
        (result, out), _ = self.apply(wb)
        self.assertEqual(kb.parts_of(out).get("chart"), 1)
        self.assertTrue(any(w["kind"] == "saved again" and "chart" in w["text"] for w in result["warnings"]))
        self.assertIn("1 chart", kb.describe(save(wb), "o.xlsx")["features"])

    def test_something_the_copy_cannot_keep_is_warned_about(self):
        data = with_sparkline(save(orders()))
        self.assertEqual(kb.parts_of(data).get("sparkline group"), 1)
        (result, out), _ = self.apply(data)
        lost = [w["text"] for w in result["warnings"] if w["kind"] == "lost"]
        self.assertEqual(lost, ["1 sparkline group in the original will not be in the copy."])

    def test_features_openpyxl_removes_are_warned_about(self):
        ext = ('<extLst><ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
               'uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}"><x14:dataValidations xmlns:xm="http://schemas.microsoft.com/office/'
               'excel/2006/main" count="1"><x14:dataValidation type="list"><x14:formula1><xm:f>Orders!$A$2:$A$3</xm:f>'
               '</x14:formula1><xm:sqref>F2</xm:sqref></x14:dataValidation></x14:dataValidations></ext></extLst>')
        data = rewrite(save(orders()), {"xl/worksheets/sheet1.xml": lambda xml: xml.replace("</worksheet>", ext + "</worksheet>")})
        (result, out), _ = self.apply(data)
        lost = [w["text"] for w in result["warnings"] if w["kind"] == "lost"]
        self.assertEqual(lost, ["Data Validation features of the original (Excel 2010 and later) will not be in the copy."])

    def test_a_macro_workbook_keeps_its_macros(self):
        data = as_xlsm(save(orders()))
        (result, out), _ = self.apply(data, name="orders.xlsm")
        self.assertTrue(result["ok"], result["message"])
        self.assertEqual(result["ext"], "xlsm")
        with zipfile.ZipFile(io.BytesIO(out)) as z:
            self.assertIn("xl/vbaProject.bin", z.namelist())
            self.assertEqual(z.read("xl/vbaProject.bin"), b"not really a macro")
            self.assertIn("macroEnabled", z.read("[Content_Types].xml").decode())

    def test_errors_times_and_dates_load_as_the_interpreter_holds_them(self):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Mixed"
        ws["A1"] = "#N/A"
        ws["A2"] = dt.datetime(2026, 3, 4, 13, 30)
        ws["A3"] = dt.datetime(2026, 3, 4)
        ws["A4"] = dt.time(6, 0)
        ws["A5"] = True
        ws["A6"] = "=text"
        ws["A6"].data_type = "s"
        sh = kb.Loaded(save(wb), "m.xlsx").sheets["Mixed"]
        self.assertEqual(sh.get(1, 1), S.XlError("#N/A"))
        self.assertEqual(sh.get(2, 1), dt.datetime(2026, 3, 4, 13, 30))
        self.assertEqual(sh.get(3, 1), dt.date(2026, 3, 4))
        self.assertEqual(sh.get(4, 1), 0.25)
        self.assertIs(sh.get(5, 1), True)
        self.assertEqual(sh.get(6, 1), "=text")

    def test_text_that_looks_like_a_formula_stays_text(self):
        wb = orders()
        wb["Orders"]["C3"] = "=not a formula"
        wb["Orders"]["C3"].data_type = "s"
        (result, out), _ = self.apply(wb)
        ws = openpyxl.load_workbook(io.BytesIO(out))["Orders"]
        hits = [c for row in ws.iter_rows() for c in row if c.value == "=not a formula"]
        self.assertEqual([(c.data_type) for c in hits], ["s"] if hits else [])

    def test_the_files_kronos_refuses(self):
        for name, data, words in (
                ("old.xls", b"\xd0\xcf\x11\xe0" + b"0" * 20, ".xls"),
                ("locked.xlsx", b"\xd0\xcf\x11\xe0" + b"0" * 20, "password"),
                ("notes.xlsx", b"just some text", "doesn't look like"),
                ("broken.xlsx", b"PK\x03\x04 broken", "doesn't look like"),
                ("other.xlsx", zipped({"hello.txt": b"hi"}), "not an Excel workbook")):
            result, out = kb.apply(EXAMPLE, data, name, {"ws": "Orders"}, TODAY)
            self.assertIsNone(out, name)
            self.assertFalse(result["ok"])
            self.assertIn(words, result["message"], name)

    def test_the_recipes_kronos_refuses(self):
        data = save(orders())
        result, out = kb.apply(EXAMPLE.replace("Cancelled", "Cancelled"), data, "o.xlsx", {"ws": "Nope"}, TODAY)
        self.assertIsNone(out)
        self.assertIn("no sheet named", result["message"])
        result, out = kb.apply(EXAMPLE, data, "o.xlsx", {}, TODAY)
        self.assertIn("Choose a sheet", result["message"])
        result, out = kb.apply("VCL 0.6\nnot a recipe", data, "o.xlsx", {}, TODAY)
        self.assertIsNone(out)
        self.assertIn("isn't proven", result["message"])
        files = (ROOT / "vcl/examples/example3_files.vcl").read_text(encoding="utf-8")
        result, out = kb.apply(files, data, "o.xlsx", {}, TODAY)
        self.assertIsNone(out)
        self.assertIn("files or folders", result["message"])

    def test_a_workbook_that_does_not_fit_the_recipe_gives_no_copy(self):
        wb = openpyxl.Workbook()
        wb.active.title = "Orders"
        wb.active.append(["Number", "Day"])
        wb.active.append([1, "x"])
        (result, out), _ = self.apply(wb)
        self.assertIsNone(out)
        self.assertTrue(result["pre_failed"])
        self.assertIn("doesn't match", result["message"])

    def test_arguments_come_with_a_default_from_the_first_test(self):
        self.assertEqual(kb.arguments(EXAMPLE), [{"name": "ws", "type": "Worksheet", "default": "Orders", "reads": False}])

    def test_describe_gives_the_ai_the_headings_and_a_sample(self):
        d = kb.describe(save(orders(30)), "orders.xlsx")
        self.assertEqual(d["sheets"][0]["headings"], ["Order No", "Order Date", "Customer", "Status"])
        self.assertEqual(d["sheets"][0]["rows"], 31)
        self.assertEqual(len(d["sheets"][0]["sample"]), kb.SAMPLE_ROWS)
        self.assertIn('Sheet "Orders": 31 rows', d["text"])
        self.assertIn("The first 10 data rows:", d["text"])
        self.assertIn("Cust 1 | Cancelled", d["text"])

    def test_a_big_sheet_comes_with_a_note_about_time(self):
        d = kb.describe(save(orders(5200)), "big.xlsx")
        self.assertEqual(d["notes"], ["The sheet Orders has 5201 rows. A recipe that deletes many rows can take a few minutes "
                                      "on a sheet this size; Cancel stops it."])
        self.assertEqual(kb.describe(save(orders(30)), "small.xlsx")["notes"], [])

    def test_the_toolchain_is_unchanged_without_an_upload(self):
        self.assertIsNone(Sheet("x").origins)
        self.assertIsNone(Sheet("x").snapshot().origins)


@unittest.skipIf(openpyxl is None, "openpyxl is not installed")
class Bridge(unittest.TestCase):
    """compiler/kronos_web.py's three calls for the page: plain JSON text, the same in CPython and Pyodide."""

    def test_describe_arguments_and_apply(self):
        import json
        import kronos_web
        data = save(orders())
        d = json.loads(kronos_web.describe_book(data, "orders.xlsx"))
        self.assertTrue(d["ok"])
        self.assertEqual(d["sheets"][0]["name"], "Orders")
        a = json.loads(kronos_web.arguments(EXAMPLE))
        self.assertEqual((a["ok"], a["files"], a["args"][0]["default"]), (True, False, "Orders"))
        self.assertTrue(json.loads(kronos_web.arguments((ROOT / "vcl/examples/example3_files.vcl").read_text(encoding="utf-8")))["files"])
        text, out = kronos_web.apply_changes(EXAMPLE, data, "orders.xlsx", json.dumps({"ws": "Orders"}), "2026-11-14")
        self.assertTrue(json.loads(text)["ok"])
        self.assertTrue(out.startswith(b"PK"))
        self.assertEqual(text, kronos_web.canonical(json.loads(text)))   # canonical: sorted, no spaces

    def test_a_file_it_cannot_read_is_a_message_not_a_crash(self):
        import json
        import kronos_web
        d = json.loads(kronos_web.describe_book(b"not a workbook", "notes.xlsx"))
        self.assertFalse(d["ok"])
        self.assertIn("doesn't look like", d["message"])
        text, out = kronos_web.apply_changes(EXAMPLE, b"not a workbook", "notes.xlsx", "{}")
        self.assertEqual(out, b"")
        self.assertFalse(json.loads(text)["ok"])


class ReferenceFormulas(unittest.TestCase):
    @unittest.skipIf(openpyxl is None, "openpyxl is not installed")
    def test_r1c1_becomes_a1(self):
        for text, row, col, want in (
                ("=RC[-1]*2", 5, 3, "=B5*2"),
                ("=R[-1]C+R2C3", 4, 2, "=B3+$C$2"),
                ("=SUM(R[-3]C:R[-1]C)", 6, 4, "=SUM(D3:D5)"),
                ('=IF(RC1="R1C1",1,0)', 2, 2, '=IF($A2="R1C1",1,0)'),
                ("=ROUND(RC[1],0)", 2, 2, "=ROUND(C2,0)")):
            self.assertEqual(kb.r1c1_to_a1(text, row, col), want)
        with self.assertRaises(kb.BookError):
            kb.r1c1_to_a1("=SUM(R:R)", 2, 2)
        with self.assertRaises(kb.BookError):
            kb.r1c1_to_a1("=R[-5]C", 2, 2)


# ---------------------------------------------------------------- workbook surgery for the tests

def zipped(parts):
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        for name, body in parts.items():
            z.writestr(name, body)
    return out.getvalue()


def rewrite(data, changes):
    """The zip with some parts replaced: changes maps a part name to a function from its text to the new text, or to
    bytes to add or replace it."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        done = set()
        for item in zin.infolist():
            body = zin.read(item.filename)
            fix = changes.get(item.filename)
            if callable(fix):
                body = fix(body.decode("utf-8")).encode("utf-8")
                done.add(item.filename)
            elif fix is not None:
                body = fix
                done.add(item.filename)
            zout.writestr(item.filename, body)
        for name, body in changes.items():
            if name not in done and isinstance(body, bytes):
                zout.writestr(name, body)
    return out.getvalue()


def with_cached_values(data, part, values):
    """Give formula cells the saved result Excel would have written (openpyxl writes none)."""
    def fix(xml):
        for ref, v in values.items():
            if v is None:
                continue
            xml, n = re.subn(rf'(<c r="{ref}"[^>]*>\s*<f>[^<]*</f>)\s*<v\s*/>|(<c r="{ref}"[^>]*>\s*<f>[^<]*</f>)\s*<v></v>',
                             lambda m: f"{m.group(1) or m.group(2)}<v>{v}</v>", xml)
            assert n == 1, (ref, xml[:600])
        return xml
    return rewrite(data, {part: fix})


def with_sparkline(data):
    ext = ('<extLst><ext xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
           'uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/'
           'excel/2006/main"><x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>Orders!A2:A5</xm:f><xm:sqref>F2'
           '</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>')
    return rewrite(data, {"xl/worksheets/sheet1.xml": lambda xml: xml.replace("</worksheet>", ext + "</worksheet>")})


def as_xlsm(data):
    def types(xml):
        xml = xml.replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                          "application/vnd.ms-excel.sheet.macroEnabled.main+xml")
        return xml.replace("</Types>", '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>')

    def rels(xml):
        return xml.replace("</Relationships>", '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/'
                           'relationships/vbaProject" Target="vbaProject.bin"/></Relationships>')
    return rewrite(data, {"[Content_Types].xml": types, "xl/_rels/workbook.xml.rels": rels,
                          "xl/vbaProject.bin": b"not really a macro"})


if __name__ == "__main__":
    unittest.main()
