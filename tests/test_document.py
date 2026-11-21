"""VCL 0.7 part 3f (D115): DOCUMENT, a spec that fills a Word template (KRONOS Offline only).

compiler/vcl_docx.py fills the template with the standard library: fields (also split over runs, in headers and
footers), a repeated row, column or section, a repeat inside a section matched to its row, an optional section; it
keeps every run's formatting and copies every other part of the file unchanged. compiler/vcl_document.py checks a
DOCUMENT against its template; the interpreter runs its TESTs on the filled document; the compiler refuses it.
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

from vba_check import check_text  # noqa: E402
from validator import ai_texts, validate_text  # noqa: E402
from vcl_codegen import DOCUMENT_ONLY, CompileError, compile_text  # noqa: E402
from vcl_docx import DocxError, Plan, Repeat, fill, template_fields  # noqa: E402
from vcl_fmt import format_text  # noqa: E402
from vcl_interp import run_tests  # noqa: E402
from vcl_mutate import mutate_report  # noqa: E402

TODAY = dt.date(2026, 11, 20)
TEMPLATES = ROOT / "data" / "templates"
ESTIMATE = (ROOT / "tests" / "data" / "doc_estimate.vcl").read_text(encoding="utf-8")
PROPOSAL = (ROOT / "tests" / "data" / "doc_proposal.vcl").read_text(encoding="utf-8")


def parts(data):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return {n: z.read(n) for n in z.namelist()}


def errors(text):
    return [p.message for p in validate_text(text) if p.severity == "error"]


class TheTemplates(unittest.TestCase):
    def test_what_the_reader_finds(self):
        est = template_fields((TEMPLATES / "estimate_template.docx").read_bytes())
        self.assertEqual(est["order"], ["Assumption", "Cost"])
        self.assertEqual(est["header"], {"Est_No", "Date"})
        self.assertEqual(est["table_section"], [None, None, None, "Cost", None])
        self.assertIn("End_User", est["body"])   # split over two runs in the template
        prop = template_fields((TEMPLATES / "proposal_template.docx").read_bytes())
        self.assertEqual(prop["table_section"], [None, None, "Equipment", "Equipment", "Nace", None])

    def test_what_the_reader_refuses(self):
        with self.assertRaisesRegex(DocxError, "not a Word"):
            template_fields(b"plain text")


class Filling(unittest.TestCase):
    def setUp(self):
        self.data = (TEMPLATES / "estimate_template.docx").read_bytes()
        self.plan = Plan(fields={"Est_No": "EST-1", "Date": "20 Nov 2026", "Enquiry": "E-1", "Title": "T",
                                 "Client": "C", "End_User": "Demo End User", "Price_Basis": "P", "Total_Qty": "2",
                                 "Grand_Total": "1,00,000"},
                         repeats=[Repeat("row", 2, [["1", "A", "2", "10", "50,000", "1,00,000"]],
                                         {"Sl": 1, "Item": 2, "Qty": 3, "Weight": 4, "Unit_Price": 5, "Total": 6}),
                                  Repeat("section", "Assumption", [["one"], ["two"], ["three"]], {"Text": 1}),
                                  Repeat("column", 3, [["T-1", "1", "2", "3"], ["T-2", "4", "5", "6"]],
                                         {"Tag": 1, "Shell_Wt": 2, "Ends_Wt": 3, "Fin_Wt": 4}),
                                  Repeat("section", "Cost", [], {"Sl": 1, "Item": 2, "Materials": 3, "Conversion": 4,
                                                                 "Factory_Cost": 5, "Unit_Price": 6})])

    def test_fields_keep_their_runs_formatting_and_everything_else_is_copied(self):
        out, report = fill(self.data, self.plan)
        before, after = parts(self.data), parts(out)
        self.assertEqual(set(before), set(after))
        for name in before:
            if name not in ("word/document.xml", "word/header1.xml", "word/footer1.xml"):
                self.assertEqual(before[name], after[name], name)
        doc = after["word/document.xml"].decode("utf-8")
        # the title run is bold and large in the template; its filled text keeps both
        self.assertRegex(doc, r"<w:r><w:rPr><w:b ?/><w:sz w:val=\"28\" ?/></w:rPr><w:t[^>]*>ESTIMATION SHEET - "
                              r"ENQUIRY E-1</w:t></w:r>")
        self.assertIn("Demo End User", report.text)                    # the field Word split over two runs
        self.assertIn(b"EST-1", after["word/header1.xml"])           # a field in the header
        self.assertIn(b"NUMPAGES", after["word/footer1.xml"])        # the page fields stay
        self.assertIn(b"rIdLogo", after["word/header1.xml"])         # and the logo
        self.assertEqual(report.left_over, 0)
        self.assertEqual(report.sections, {"Assumption": 3, "Cost": 0})
        self.assertEqual(report.tables[2], [["Parameter", "T-1", "T-2"], ["Shell", "1", "4"],
                                            ["Dished ends / flanges", "2", "5"], ["Finished weight per item", "3", "6"]])
        self.assertNotIn("{{", report.text)

    def test_namespaces_the_template_declares_are_kept(self):
        """Word lists prefixes in mc:Ignorable; a prefix nothing else uses must still be declared after filling."""
        src = parts(self.data)
        doc = src["word/document.xml"].decode("utf-8").replace(
            "<w:document ", '<w:document xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
            'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14" ', 1)
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as z:
            for n, b in src.items():
                z.writestr(n, doc.encode("utf-8") if n == "word/document.xml" else b)
        out, _ = fill(buf.getvalue(), self.plan)
        head = parts(out)["word/document.xml"].decode("utf-8").split(">", 2)[1]
        self.assertIn('xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"', head)
        self.assertIn('Ignorable="w14"', head)

    def test_a_row_value_beyond_the_source_row_is_refused(self):
        self.plan.repeats[0] = Repeat("row", 2, [["1", "A"]], {"Sl": 1, "Item": 2, "Total": 6})
        with self.assertRaisesRegex(DocxError, "Total=6"):
            fill(self.data, self.plan)


class TheSpecs(unittest.TestCase):
    def test_both_practice_documents_are_valid_canonical_and_pass(self):
        for text in (ESTIMATE, PROPOSAL):
            self.assertEqual(errors(text), [])
            self.assertEqual(format_text(text), text)
            results = run_tests(text, TODAY)
            self.assertTrue(results and all(r.passed for r in results), [r.failed for r in results])
            self.assertTrue(all(r.document and r.document.startswith(b"PK") for r in results))

    def test_kronos_offline_only(self):
        warnings = [p.message for p in validate_text(PROPOSAL) if p.severity == "warning"]
        self.assertTrue(any(w.startswith("KRONOS Offline only") for w in warnings))
        with self.assertRaises(CompileError) as caught:
            compile_text(PROPOSAL, "doc_proposal")
        self.assertEqual(str(caught.exception), DOCUMENT_ONLY)
        self.assertIn("KRONOS Offline", check_text(PROPOSAL, "doc_proposal", TODAY).skipped)

    def test_its_ai_texts_are_listed(self):
        self.assertEqual([t for _, t in ai_texts(PROPOSAL)],
                         ["Offer based on the enquiry data sheet and ASME VIII-1, latest edition.",
                          "Two gussets on every nozzle.", "Repads as thick as the shell."])

    def test_the_tests_catch_deliberate_bugs(self):
        caught, total, survivors = mutate_report(PROPOSAL, snapshot=True, today=TODAY)
        self.assertEqual((caught, total), (total, total), survivors)
        self.assertGreater(total, 50)

    def test_what_the_validator_refuses(self):
        cases = [
            (PROPOSAL.replace("VCL 0.7", "VCL 0.6", 1), "DOCUMENT is VCL 0.7"),
            (PROPOSAL.replace('FIELD    Total_Qty = "4 Nos"\n', ""), "{{Total_Qty}} has no FIELD line"),
            (PROPOSAL.replace('FIELD    Total_Qty = "4 Nos"', 'FIELD    Total_Qty = "4 Nos"\nFIELD    Weight = "1"'),
             "FIELD Weight: the template has no {{Weight}}"),
            (PROPOSAL.replace("KEEP     Nace IF $sour\n", ""), "section Nace needs a REPEAT section"),
            (PROPOSAL.replace(" match=1=Tag", ""), "say which rows belong to each copy"),
            (PROPOSAL.replace("table=6 from", "table=9 from"), "has no table 9"),
            (PROPOSAL.replace(", Nozzle_Notes=9", ""), "field Nozzle_Notes is not bound"),
            (PROPOSAL.replace("ARGS     proposalNo As String", "ARGS     proposalNo As Range"), "a DOCUMENT takes only"),
            (PROPOSAL.replace('FIELD    Client = "Demo Compressors Ltd"', "FIELD    Client = client"),
             "a DOCUMENT has no variables"),
            (PROPOSAL.replace('FIELD    Client = "Demo Compressors Ltd"', 'FIELD    Client = NUMBER_TEXT(5, "0.000")'),
             "NUMBER_TEXT needs one of"),
            (PROPOSAL.replace("SECTION_COUNT(Equipment) = 2", "SECTION_COUNT(Equipment) = 2 AND VALUE(Notes!A2) = 1"),
             "VALUE is not one of"),
            (PROPOSAL.replace("FIELD_TEXT(Proposal_No)", "FIELD_TEXT(Mark)"), "FIELD_TEXT(Mark): Mark has no FIELD"),
        ]
        for text, want in cases:
            with self.subTest(want=want):
                found = errors(text)
                self.assertTrue(any(want in m for m in found), found)

    def test_a_filled_value_that_is_not_a_number_stops_the_test(self):
        bad = ESTIMATE.replace('FIELD    Title = "Channel and floating head assemblies, 2 sets"',
                               'FIELD    Title = NUMBER_TEXT(CELL(SHEET($wb, "Enq Summary"), 2, 10), "0")')
        results = run_tests(bad, TODAY)
        self.assertFalse(results[0].passed)
        self.assertTrue(any(re.search("could not be filled", m) for m in results[0].failed), results[0].failed)


if __name__ == "__main__":
    unittest.main()
