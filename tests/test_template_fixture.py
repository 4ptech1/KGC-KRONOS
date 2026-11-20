"""VCL 0.7 part 3a (D110): template FIXTUREs, `FIXTURE <file> template=<name>`.

The reader (compiler/vcl_template.py) turns a look-alike template of data/templates/ into the practice workbook the
interpreter, the compiler's fixture builder and the VBA simulator all use: values, formulas with their saved values,
number formats, fills and merged ranges. It refuses what it can't hold. A practice spec on the WPS template passes
every stage, and every 0.6 spec stays valid under the VCL 0.7 line (E16).
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

from validator import validate_text  # noqa: E402
from vba_check import check_text  # noqa: E402
from vcl_codegen import compile_text  # noqa: E402
from vcl_fixtures import Formula, fixtures_module  # noqa: E402
from vcl_interp import run_tests  # noqa: E402
from vcl_parse import parse  # noqa: E402
from vcl_template import TemplateError, load, load_template  # noqa: E402

try:
    import openpyxl
except ImportError:   # CI installs requirements.txt
    openpyxl = None

TODAY = dt.date(2026, 11, 20)
PRACTICE = (ROOT / "tests" / "data" / "template_wps.vcl").read_text(encoding="utf-8")
TEMPLATES = ROOT / "data" / "templates"


def errors(text):
    return [p for p in validate_text(text) if p.severity == "error"]


def rezip(data, change):
    """The same .xlsx with one part changed: change(name, bytes) -> bytes."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as src, zipfile.ZipFile(out, "w") as dst:
        for name in src.namelist():
            dst.writestr(name, change(name, src.read(name)))
    return out.getvalue()


class TheReader(unittest.TestCase):
    def test_it_reads_the_wps_form(self):
        sheets = load_template("wps_form")
        self.assertEqual(list(sheets), ["WPS"])
        ws = sheets["WPS"]
        self.assertEqual(ws.get(6, 1), "Welding Procedure Specification No. :")
        self.assertEqual(ws.get(25, 3), 12)
        self.assertIn((6, 3, 6, 5), ws.merged)     # C6:E6
        self.assertEqual(len(ws.merged), 97)
        f = ws.get(29, 18)                          # R29, the heat input
        self.assertIsInstance(f, Formula)
        self.assertEqual(f.text, "=$N$25*N29*L29*60/(Q29*1000)")
        self.assertAlmostEqual(f.saved, 0.36)
        self.assertEqual(ws.get(31, 20).saved, "OVER LIMIT")
        self.assertEqual(ws.formats[(29, 18)], "0.000")
        self.assertEqual(ws.fills[(25, 14)], "#FFF2CC")

    def test_it_keeps_the_sheet_order_and_names_of_the_estimation_workbook(self):
        self.assertEqual(list(load_template("estimation")), ["Summary", "Inputs", "BOM-RM", "Conversion",
                                                            "Hydrotest Options", "Enq Summary",
                                                            "Assumptions & Clarif."])

    @unittest.skipIf(openpyxl is None, "openpyxl is not installed")
    def test_every_value_and_saved_value_is_what_openpyxl_reads(self):
        for name in ("wps_form", "estimation", "enquiry_register"):
            path = TEMPLATES / f"{name}.xlsx"
            formulas, values = openpyxl.load_workbook(path), openpyxl.load_workbook(path, data_only=True)
            for title, sh in load_template(name).items():
                ws, vs = formulas[title], values[title]
                seen = {(c.row, c.column) for row in ws.iter_rows() for c in row if c.value not in (None, "")}
                self.assertEqual(set(sh.cells), seen, f"{name} {title}")
                for (r, c), v in sh.cells.items():
                    with self.subTest(cell=f"{name} {title}!{ws.cell(r, c).coordinate}"):
                        if isinstance(v, Formula):
                            self.assertEqual(v.text, ws.cell(r, c).value)
                            self.assertEqual(v.saved, vs.cell(r, c).value)
                        else:
                            self.assertEqual(v, ws.cell(r, c).value)
                merged = {(m.min_row, m.min_col, m.max_row, m.max_col) for m in ws.merged_cells.ranges}
                self.assertEqual(set(sh.merged), merged)

    @unittest.skipIf(openpyxl is None, "openpyxl is not installed")
    def test_a_number_with_a_date_format_is_a_date(self):
        wb = openpyxl.Workbook()
        wb.active["A1"] = dt.date(2026, 11, 20)
        wb.active["A1"].number_format = "dd/mm/yyyy"
        wb.active["B1"] = 46000
        out = io.BytesIO()
        wb.save(out)
        sh = load(out.getvalue())["Sheet"]
        self.assertEqual(sh.get(1, 1), dt.date(2026, 11, 20))
        self.assertEqual(sh.get(1, 2), 46000)

    def test_what_it_refuses(self):
        data = (TEMPLATES / "wps_form.xlsx").read_bytes()
        shared = rezip(data, lambda n, b: re.sub(rb"<f>(L29&amp;)", rb'<f t="shared" ref="D29" si="0">\1', b, 1)
                       if n == "xl/worksheets/sheet1.xml" else b)
        with self.assertRaisesRegex(TemplateError, "shared formula"):
            load(shared, "x.xlsx")
        error = rezip(data, lambda n, b: b.replace(b'<c r="C25" t="n"><v>12</v></c>',
                                                   b'<c r="C25" t="e"><v>#N/A</v></c>', 1)
                      if n == "xl/worksheets/sheet1.xml" else b)
        with self.assertRaisesRegex(TemplateError, "error value"):
            load(error, "x.xlsx")
        with self.assertRaisesRegex(TemplateError, "not an .xlsx"):
            load(b"not a zip", "x.xlsx")
        with self.assertRaisesRegex(TemplateError, "no data/templates/nothing_here.xlsx"):
            load_template("nothing_here")
        with self.assertRaisesRegex(TemplateError, "letters, digits and _"):
            load_template("../secret")


class TheSpecLine(unittest.TestCase):
    def test_the_practice_spec_passes_every_stage(self):
        self.assertEqual(errors(PRACTICE), [])
        results = run_tests(PRACTICE, TODAY)
        self.assertTrue(results and all(r.passed for r in results), [r.failed for r in results])
        self.assertIn("Function SetWpsRevision", compile_text(PRACTICE, "template_wps"))
        outcome = check_text(PRACTICE, "template_wps", TODAY)
        self.assertTrue(outcome.proven, outcome.differences or outcome.skipped)

    def test_the_excel_builder_writes_values_formulas_formats_fills_and_merges(self):
        spec, _ = parse(PRACTICE)
        trees = [k["tree"] for k in spec["keys"] if k["key"] == "FIXTURE"]
        vba = fixtures_module([("template_wps", trees)])
        self.assertIn("Public Sub ForgeFixture_wps_01(ByVal wb As Workbook)", vba)
        self.assertIn('ws.Range("R29").Formula = "=$N$25*N29*L29*60/(Q29*1000)"', vba)
        self.assertIn('ws.Range("C6:E6").Merge', vba)
        self.assertIn('ws.Range("R29").NumberFormat = "0.000"', vba)
        self.assertIn('ws.Range("N25").Interior.Color = RGB(255, 242, 204)', vba)
        self.assertIn('ws.Range("C25").Value = 12', vba)

    def test_template_fixtures_are_vcl_07(self):
        old = PRACTICE.replace("VCL 0.7", "VCL 0.6", 1)
        self.assertIn("V333", " ".join(p.code + " " + p.message for p in errors(old)))

    def test_one_template_line_per_file_and_no_sheet_lines_beside_it(self):
        two = PRACTICE.replace("FIXTURE  wps_01.xlsx template=wps_form",
                               "FIXTURE  wps_01.xlsx template=wps_form\nFIXTURE  wps_01.xlsx template=wps_form")
        self.assertIn("V334", " ".join(p.code + " " + p.message for p in errors(two)))
        mixed = PRACTICE.replace("FIXTURE  wps_01.xlsx template=wps_form",
                                 "FIXTURE  wps_01.xlsx template=wps_form\n"
                                 "FIXTURE  wps_01.xlsx sheet=Notes cols=A:Note rows=1 data=A:values(\"x\"x1)")
        self.assertIn("V335", " ".join(p.code + " " + p.message for p in errors(mixed)))

    def test_a_missing_template_is_named(self):
        bad = PRACTICE.replace("template=wps_form", "template=no_such_form")
        self.assertIn("no data/templates/no_such_form.xlsx", " ".join(p.code + " " + p.message for p in errors(bad)))

    def test_every_06_spec_stays_valid_and_gives_the_same_results_as_07(self):
        for path in sorted((ROOT / "data" / "verified").glob("c*/spec.vcl"))[::10]:   # every tenth 0.6 pair
            text = path.read_text(encoding="utf-8")
            if not text.startswith("VCL 0.6"):
                continue
            with self.subTest(spec=path.parent.name):
                newer = text.replace("VCL 0.6", "VCL 0.7")   # every TASK of a BUILD too
                self.assertEqual(errors(newer), [])
                self.assertEqual([(r.passed, r.result) for r in run_tests(newer, TODAY)],
                                 [(r.passed, r.result) for r in run_tests(text, TODAY)])


if __name__ == "__main__":
    unittest.main()
