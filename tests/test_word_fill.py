"""KRONOS Offline version 6, part 6d (D121): a DOCUMENT recipe fills the person's own Word template.

kronos_book.fill_document proves the recipe as KRONOS Offline proves it (D118), reads the workbook it takes (if any) as
Excel last saved it, and fills the uploaded .docx with the interpreter's meaning of its FIELD, REPEAT and KEEP lines
(vcl_interp.doc_plan, vcl_docx.fill): the very file its TESTs check. A formula with no saved value stops the fill in
words, never guessed. Neither the template nor the workbook is ever changed.
"""
import datetime as dt
import io
import json
import sys
import unittest
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(ROOT / "vcl"))

import kronos_book as kb  # noqa: E402
import kronos_web  # noqa: E402
from vcl_interp import run_tests  # noqa: E402

TODAY = dt.date(2026, 11, 20)
TEMPLATES = ROOT / "data" / "templates"
PROPOSAL = (ROOT / "vcl" / "examples" / "example5_document.vcl").read_text(encoding="utf-8")
ESTIMATE = (ROOT / "tests" / "data" / "doc_estimate.vcl").read_text(encoding="utf-8")
PROPOSAL_DOCX = (TEMPLATES / "proposal_template.docx").read_bytes()
ESTIMATE_DOCX = (TEMPLATES / "estimate_template.docx").read_bytes()
ESTIMATION = (TEMPLATES / "estimation.xlsx").read_bytes()
TASK = (ROOT / "vcl" / "examples" / "example1_rev3.vcl").read_text(encoding="utf-8")


def with_extra_field(data):
    """The proposal template with one more {{field}} that no recipe fills."""
    src = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n in src.namelist():
            body = src.read(n)
            if n == "word/document.xml":
                body = body.replace(b"<w:body>", b"<w:body><w:p><w:r><w:t>{{Not_In_The_Recipe}}</w:t></w:r></w:p>", 1)
            z.writestr(n, body)
    return out.getvalue()


class FromTypedValues(unittest.TestCase):
    def test_the_proposal_is_the_one_its_test_checks(self):
        result, out = kb.fill_document(PROPOSAL, PROPOSAL_DOCX, "Proposal.docx",
                                       {"proposalNo": "P-9101-R0", "onDate": "20 Nov 2026", "sour": "FALSE"}, TODAY)
        self.assertTrue(result["ok"], result["message"])
        self.assertEqual(out, run_tests(PROPOSAL, TODAY)[0].document)
        self.assertEqual((result["sections"], result["left_over"], result["warnings"], result["read_name"], result["stem"]),
                         ({"Equipment": 2, "Nace": 0}, 0, [], "", "Proposal"))
        self.assertEqual(dict((f["name"], f["text"]) for f in result["fields"])["Proposal_No"], "P-9101-R0")
        self.assertEqual([t["text"] for t in result["ai_texts"]],
                         ["Offer based on the enquiry data sheet and ASME VIII-1, latest edition.",
                          "Two gussets on every nozzle.", "Repads as thick as the shell."])

    def test_its_typed_values_are_the_persons(self):
        result, out = kb.fill_document(PROPOSAL, PROPOSAL_DOCX, "Proposal.docx",
                                       {"proposalNo": "P-9101-R1", "onDate": "21 Nov 2026", "sour": "TRUE"}, TODAY)
        self.assertEqual(out, run_tests(PROPOSAL, TODAY)[1].document)
        self.assertEqual(result["sections"]["Nace"], 1)
        result, _ = kb.fill_document(PROPOSAL, PROPOSAL_DOCX, "Proposal.docx", {"sour": "maybe"}, TODAY)
        self.assertEqual(result["message"], "sour needs TRUE or FALSE.")

    def test_a_field_nothing_fills_is_warned_about(self):
        result, out = kb.fill_document(PROPOSAL, with_extra_field(PROPOSAL_DOCX), "Proposal.docx", {"sour": "FALSE"}, TODAY)
        self.assertTrue(result["ok"])
        self.assertEqual(result["left_over"], 1)
        self.assertEqual(result["warnings"], [{"kind": "check", "text": "1 field in the template was not filled: look for "
                                                                         "{{ in the document."}])

    def test_the_template_is_never_changed(self):
        template = bytes(PROPOSAL_DOCX)
        kb.fill_document(PROPOSAL, template, "Proposal.docx", {"sour": "FALSE"}, TODAY)
        self.assertEqual(template, PROPOSAL_DOCX)


class FromAWorkbook(unittest.TestCase):
    def test_the_estimation_sheet_is_the_one_its_test_checks(self):
        result, out = kb.fill_document(ESTIMATE, ESTIMATE_DOCX, "Estimate.docx", {"estNo": "EST-9001", "onDate": "20 Nov 2026"},
                                       TODAY, book=(ESTIMATION, "E-9100.xlsx"))
        self.assertTrue(result["ok"], result["message"])
        self.assertEqual(out, run_tests(ESTIMATE, TODAY)[0].document)
        self.assertEqual(result["read_name"], "E-9100.xlsx")

    def test_without_the_workbook_nothing_is_filled(self):
        result, out = kb.fill_document(ESTIMATE, ESTIMATE_DOCX, "Estimate.docx", {}, TODAY)
        self.assertIsNone(out)
        self.assertEqual(result["message"], 'This recipe reads a workbook (wb): choose it under "Your workbook" first.')

    def test_a_formula_with_no_saved_value_stops_the_fill_in_words(self):
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(ESTIMATION))   # openpyxl saves formulas without their values
        saved = io.BytesIO()
        wb.save(saved)
        result, out = kb.fill_document(ESTIMATE, ESTIMATE_DOCX, "Estimate.docx", {}, TODAY, book=(saved.getvalue(), "e.xlsx"))
        self.assertIsNone(out)
        self.assertIn("with no saved value: open the workbook in Excel and save it", result["message"])


class Refused(unittest.TestCase):
    def test_what_kronos_will_not_fill(self):
        cases = [((PROPOSAL, PROPOSAL_DOCX, "Proposal.doc"), "This is an old .doc file. In Word choose File, Save As, and "
                  "save it as .docx; then choose that."),
                 ((PROPOSAL, PROPOSAL_DOCX, "Proposal.pdf"), "KRONOS fills Word templates saved as .docx."),
                 ((PROPOSAL, b"not a zip", "Proposal.docx"), "The template could not be filled: this is not a Word .docx file"),
                 ((TASK, PROPOSAL_DOCX, "Proposal.docx"), "This recipe changes a workbook; it fills no Word document."),
                 ((PROPOSAL.replace("TEMPLATE proposal_template.docx", "TEMPLATE nothing_here.docx"), PROPOSAL_DOCX,
                   "Proposal.docx"), "This recipe isn't proven, so KRONOS won't fill your template with it.")]
        for (text, template, name), words in cases:
            with self.subTest(name=name, words=words):
                result, out = kb.fill_document(text, template, name, {"sour": "FALSE"}, TODAY)
                self.assertIsNone(out)
                self.assertTrue(result["message"].startswith(words), result["message"])

    def test_a_template_the_recipe_does_not_fit_is_a_message_not_a_crash(self):
        result, out = kb.fill_document(PROPOSAL, ESTIMATE_DOCX, "Estimate.docx", {"sour": "FALSE"}, TODAY)
        self.assertIsNone(out)
        self.assertFalse(result["ok"])
        self.assertTrue(result["message"].startswith("The template"), result["message"])


class ThePage(unittest.TestCase):
    def test_the_bridge_gives_the_same(self):
        text, data = kronos_web.fill_document(ESTIMATE, ESTIMATE_DOCX, "Estimate.docx", json.dumps({"estNo": "X"}), TODAY,
                                              ESTIMATION, "E-9100.xlsx")
        want, out = kb.fill_document(ESTIMATE, ESTIMATE_DOCX, "Estimate.docx", {"estNo": "X"}, TODAY,
                                     book=(ESTIMATION, "E-9100.xlsx"))
        self.assertEqual(text, kronos_web.canonical(want))
        self.assertEqual(data, out)
        self.assertEqual([a["name"] for a in json.loads(kronos_web.arguments(PROPOSAL))["args"]],
                         ["proposalNo", "onDate", "sour"])


if __name__ == "__main__":
    unittest.main()
