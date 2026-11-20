"""The look-alike KGC templates (D108, data/templates/): invented practice copies of KGC's documents.

Every committed template must be what data/templates/make_templates.py makes today (the same zip parts, byte for byte,
once unzipped), and must hold what the tests and recipes of the KGC documents plan depend on: the WPS form's merged
cells, pictures, print area and heat-input check; the estimation workbook's sheets and formulas; the register's merged
enquiry blocks; the Word templates' fields, repeated parts, header and footer. The saved formula values are checked
against the rules worked out here independently.
"""
import io
import math
import re
import sys
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TEMPLATES = ROOT / "data" / "templates"
sys.path.insert(0, str(TEMPLATES))

try:
    import openpyxl
except ImportError:   # CI installs requirements.txt
    openpyxl = None

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
FIELD = re.compile(r"\{\{([A-Za-z0-9_]+)\}\}")
MARKER = re.compile(r"^\{\{(BEGIN|END) ([A-Za-z0-9_]+)\}\}$")


def parts(data):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return {n: z.read(n) for n in z.namelist()}


def paragraphs(xml):
    """The text of each paragraph, its runs read together."""
    root = ET.fromstring(xml)
    return ["".join(t.text or "" for t in p.iter(W + "t")) for p in root.iter(W + "p")]


@unittest.skipIf(openpyxl is None, "openpyxl is not installed")
class TheFilesAreWhatTheScriptMakes(unittest.TestCase):
    def test_each_template_is_what_make_templates_makes_today(self):
        import make_templates
        on_disk = sorted(p.name for p in TEMPLATES.iterdir() if p.suffix in (".xlsx", ".docx"))
        self.assertEqual(on_disk, sorted(make_templates.FILES))
        for name in make_templates.FILES:
            with self.subTest(file=name):
                self.assertEqual(parts((TEMPLATES / name).read_bytes()), parts(make_templates.build(name)),
                                 "run: python data/templates/make_templates.py")

    def test_the_readme_names_every_template(self):
        import make_templates
        readme = (TEMPLATES / "README.md").read_text(encoding="utf-8")
        for name in make_templates.FILES:
            self.assertIn(f"`{name}`", readme)


@unittest.skipIf(openpyxl is None, "openpyxl is not installed")
class TheWorkbooks(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.books = {}
        for name in ("wps_form", "estimation", "enquiry_register"):
            path = TEMPLATES / f"{name}.xlsx"
            cls.books[name] = (openpyxl.load_workbook(path), openpyxl.load_workbook(path, data_only=True))

    def test_every_formula_has_a_saved_value(self):
        for name, (wb, values) in self.books.items():
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for c in row:
                        if isinstance(c.value, str) and c.value.startswith("="):
                            with self.subTest(cell=f"{name} {ws.title}!{c.coordinate}"):
                                self.assertIsNotNone(values[ws.title][c.coordinate].value)

    def test_no_link_to_another_workbook_and_no_macro(self):
        for name in self.books:
            names = parts((TEMPLATES / f"{name}.xlsx").read_bytes())
            self.assertFalse([n for n in names if "externalLink" in n or "vbaProject" in n], name)

    def test_the_wps_form_keeps_the_forms_layout(self):
        wb, values = self.books["wps_form"]
        ws = wb["WPS"]
        self.assertEqual(wb.sheetnames, ["WPS"])
        self.assertEqual(len(ws.merged_cells.ranges), 97)
        self.assertEqual(ws.print_area, "'WPS'!$A$1:$I$50")
        self.assertEqual([b.id for b in ws.row_breaks.brk], [33])
        self.assertEqual(len(ws._images), 2)   # the logo and the weld sketch
        self.assertIn("A14:C19", {str(r) for r in ws.merged_cells.ranges})   # the sketch's place
        self.assertEqual(ws["A25"].value, "Material Thickness (mm)")

    def test_the_wps_heat_input_check_follows_its_rule(self):
        wb, values = self.books["wps_form"]
        v = values["WPS"]
        k, limit = v["N25"].value, v["N26"].value
        outcomes = []
        for r in (29, 30, 31):
            i1, i2, u1, u2, s1, s2 = (v[f"{c}{r}"].value for c in "LMNOPQ")
            lo, hi = k * u1 * i1 * 60 / (s2 * 1000), k * u2 * i2 * 60 / (s1 * 1000)
            self.assertAlmostEqual(v[f"R{r}"].value, lo)
            self.assertAlmostEqual(v[f"S{r}"].value, hi)
            self.assertEqual(v[f"T{r}"].value, "OK" if hi <= limit else "OVER LIMIT")
            self.assertEqual(v[f"D{r}"].value, f"{i1}–{i2}")
            outcomes.append((v[f"T{r}"].value, v[f"AC{r}"].value))
        # the practice data has one run over the limit and one PQR run outside the WPS range
        self.assertEqual(outcomes, [("OK", "Yes"), ("OK", "No"), ("OVER LIMIT", "Yes")])

    def test_the_estimation_workbook_has_the_seven_sheets_and_their_rules(self):
        wb, values = self.books["estimation"]
        self.assertEqual(wb.sheetnames, ["Summary", "Inputs", "BOM-RM", "Conversion", "Hydrotest Options",
                                         "Enq Summary", "Assumptions & Clarif."])
        f = wb["BOM-RM"]
        self.assertTrue(f["M5"].value.startswith("=PI()/4*"))   # a ring
        self.assertTrue(f["M6"].value.startswith("=(I6+2*"))    # a plate
        self.assertEqual(f["L10"].value, '=SUMIF($A$5:$A$8,"CH",L5:L8)')
        s = values["Summary"]
        for ref in ("B17", "C17", "B18", "C18"):
            price = s[ref].value
            self.assertEqual(price % 1000, 0, ref)
            self.assertTrue(wb["Summary"][ref].value.startswith("=ROUNDUP("), ref)
        margin = values["Inputs"]["C15"].value
        self.assertEqual(s["B17"].value, math.ceil(s["B14"].value * (1 + margin) / 2 / 1000) * 1000)
        e = values["Enq Summary"]
        self.assertEqual(e["Q2"].value, e["R2"].value + 2 * e["S2"].value)
        self.assertEqual(e["AB2"].value, s["B17"].value)

    def test_the_register_merges_each_enquiry_over_its_items(self):
        wb, values = self.books["enquiry_register"]
        ws = wb.active
        merged = {str(r) for r in ws.merged_cells.ranges}
        self.assertEqual(merged, {f"{c}{a}:{c}{b}" for c in "ABCDEFGHI" for a, b in ((2, 3), (4, 6))})
        v = values[ws.title]
        for r in range(2, 8):
            self.assertEqual(v[f"Q{r}"].value, v[f"R{r}"].value + 2 * v[f"S{r}"].value)
            barg = float(re.search(r"ROUND\(([\d.]+)\*", ws[f"W{r}"].value).group(1))
            self.assertAlmostEqual(v[f"W{r}"].value, round(barg * 1.01972, 2))
        self.assertEqual(ws["C2"].value, "05.10.2026")   # dates are text
        self.assertEqual(sorted(c.coordinate for row in ws.iter_rows() for c in row if c.comment), ["B2", "W5"])


class TheWordTemplates(unittest.TestCase):
    EXPECTED = {
        "estimate_template.docx": {
            "fields": {"Enquiry", "Title", "Client", "End_User", "Price_Basis", "Sl", "Item", "Qty", "Weight",
                       "Unit_Price", "Total", "Total_Qty", "Grand_Total", "Text", "Tag", "Shell_Wt", "Ends_Wt",
                       "Fin_Wt", "Materials", "Conversion", "Factory_Cost"},
            "header": {"Est_No", "Date"},
            "sections": ["Assumption", "Cost"],
        },
        "proposal_template.docx": {
            "fields": {"Client", "Item_Line", "End_User", "Enquiry_Ref", "Proposal_No", "Date", "Basis", "Sl",
                       "Item", "Tag", "Qty", "Total_Qty", "Shell_ID", "TL_TL", "Thk", "Design_Pr", "Mark", "Nozzle",
                       "Size", "Rating", "Nozzle_Notes", "Requirement", "Confirmation", "TQ_No", "Query"},
            "header": {"Proposal_No", "Date"},
            "sections": ["Equipment", "Nace"],
        },
    }

    def test_fields_markers_header_and_footer(self):
        for name, want in self.EXPECTED.items():
            with self.subTest(file=name):
                p = parts((TEMPLATES / name).read_bytes())
                body = paragraphs(p["word/document.xml"])
                fields = {f for text in body if not MARKER.match(text) for f in FIELD.findall(text)}
                self.assertEqual(fields, want["fields"])
                stack, sections = [], []
                for text in body:
                    m = MARKER.match(text)
                    if m and m.group(1) == "BEGIN":
                        stack.append(m.group(2))
                    elif m:
                        self.assertEqual(stack.pop(), m.group(2), "BEGIN and END must pair up")
                        sections.append(m.group(2))
                self.assertEqual((stack, sections), ([], want["sections"]))
                for text in body:   # a marker is a paragraph on its own; no stray braces
                    self.assertEqual(FIELD.sub("", text).count("{{") + FIELD.sub("", text).count("}}"),
                                     0 if not MARKER.match(text) else 2, text)
                self.assertEqual({f for t in paragraphs(p["word/header1.xml"]) for f in FIELD.findall(t)},
                                 want["header"])
                self.assertIn(b"rIdLogo", p["word/header1.xml"])
                self.assertIn("word/media/logo.png", p)
                self.assertIn(b"NUMPAGES", p["word/footer1.xml"])

    def test_one_field_is_split_over_two_runs_as_word_often_stores_it(self):
        xml = parts((TEMPLATES / "estimate_template.docx").read_bytes())["word/document.xml"].decode("utf-8")
        self.assertIn("{{End_</w:t>", xml)
        self.assertIn(">User}}</w:t>", xml)


if __name__ == "__main__":
    unittest.main()
