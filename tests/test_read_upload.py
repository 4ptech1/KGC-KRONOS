"""KRONOS Offline version 6, part 6c (D120): a recipe that reads a second workbook, uploaded.

A recipe whose TEST has read= (D114) takes a Workbook it only reads. KRONOS Offline asks for it as a second upload,
"Workbook to read from": kronos_book reads it as Excel last saved it (vcl_template's reading mode, so every formula
gives its saved value), binds it to that argument, and never changes or writes it. The other Workbook arguments, and
the copy, are the first upload's, as before.
"""
import datetime as dt
import io
import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(ROOT / "vcl"))
sys.path.insert(0, str(ROOT / "tests"))

try:
    import openpyxl
except ImportError:   # CI installs requirements.txt
    openpyxl = None

TODAY = dt.date(2026, 11, 20)
RECIPE = (ROOT / "tests" / "data" / "wps_from_wps.vcl").read_text(encoding="utf-8")
NEW = (ROOT / "data" / "templates" / "wps_form.xlsx").read_bytes()
OLD = (ROOT / "data" / "templates" / "wps_form_b.xlsx").read_bytes()
EXAMPLE = (ROOT / "vcl" / "examples" / "example1_rev3.vcl").read_text(encoding="utf-8")


class Arguments(unittest.TestCase):
    def test_the_workbook_a_recipe_only_reads_is_marked(self):
        import kronos_book as kb
        self.assertEqual(kb.arguments(RECIPE), [
            {"name": "old", "type": "Workbook", "default": "wps_old_01.xlsx", "reads": True},
            {"name": "ws", "type": "Worksheet", "default": "WPS", "reads": False}])
        self.assertTrue(kb.needs_reading(RECIPE))
        self.assertFalse(kb.needs_reading(EXAMPLE))
        self.assertFalse(any(a["reads"] for a in kb.arguments(EXAMPLE)))


@unittest.skipIf(openpyxl is None, "openpyxl is not installed")
class TwoUploads(unittest.TestCase):
    def apply(self, read=(OLD, "WPS-105.xlsx"), new=NEW):
        import kronos_book as kb
        return kb.apply(RECIPE, new, "WPS-201.xlsx", {"ws": "WPS"}, TODAY, read=read)

    def test_the_new_wps_is_filled_from_the_old_one(self):
        result, out = self.apply()
        self.assertTrue(result["ok"], result["message"])
        self.assertEqual((result["result"], result["read_name"], result["kept_everything"]), ("3", "WPS-105.xlsx", True))
        ws = openpyxl.load_workbook(io.BytesIO(out))["WPS"]
        self.assertEqual((ws["C7"].value, ws["L29"].value, ws["Q31"].value, ws["C6"].value),
                         ("DM-PQR-105", 80, 120, "DM-WPS-201"))
        self.assertEqual(ws["B47"].value, "2. Runs inside the heat-input limit in the old WPS: 3")

    def test_the_same_answer_as_the_recipes_own_test(self):
        from kronos_engine import prove
        proof = prove(RECIPE, TODAY, offline=True)
        self.assertTrue(proof.ok)
        result, _ = self.apply()
        self.assertEqual(str(proof.tests[0].result), result["result"])

    def test_without_the_second_workbook_nothing_is_done(self):
        result, out = self.apply(read=None)
        self.assertIsNone(out)
        self.assertEqual(result["message"], 'This recipe reads a second workbook (old): choose it under "Workbook to read '
                                            'from".')

    def test_the_workbook_to_read_from_is_never_changed(self):
        old = bytes(OLD)
        self.apply(read=(old, "old.xlsx"))
        self.assertEqual(old, OLD)

    def test_a_formula_with_no_saved_value_stops_it_in_words(self):
        wb = openpyxl.load_workbook(io.BytesIO(OLD))   # openpyxl saves formulas without their values
        out = io.BytesIO()
        wb.save(out)
        result, copy = self.apply(read=(out.getvalue(), "old.xlsx"))
        self.assertIsNone(copy)
        self.assertIn("open the workbook in Excel and save it", result["message"])

    def test_a_file_that_is_not_a_workbook_is_refused_in_words(self):
        result, out = self.apply(read=(b"not a workbook", "old.xlsx"))
        self.assertIsNone(out)
        self.assertIn("doesn't look like", result["message"])
        result, out = self.apply(read=(OLD, "old.xls"))
        self.assertIn(".xls", result["message"])

    def test_a_recipe_that_reads_nothing_ignores_the_second_upload(self):
        import kronos_book as kb
        from test_kronos_book import orders, save
        result, out = kb.apply(EXAMPLE, save(orders()), "orders.xlsx", {"ws": "Orders"}, TODAY, read=(OLD, "old.xlsx"))
        self.assertTrue(result["ok"])
        self.assertEqual(result["read_name"], "")

    def test_the_page_bridge_gives_the_same(self):
        import kronos_book as kb
        import kronos_web
        text, data = kronos_web.apply_changes_many(json.dumps([RECIPE]), NEW, "WPS-201.xlsx", json.dumps([{"ws": "WPS"}]),
                                                   TODAY, None, OLD, "WPS-105.xlsx")
        want, copy = kb.apply_many([RECIPE], NEW, "WPS-201.xlsx", [{"ws": "WPS"}], TODAY, read=(OLD, "WPS-105.xlsx"))
        self.assertEqual(json.loads(text), json.loads(kronos_web.canonical(want)))
        self.assertEqual(data, copy)
        self.assertTrue(json.loads(kronos_web.arguments(RECIPE))["args"][0]["reads"])

    def test_a_document_is_not_run_on_a_workbook(self):
        import kronos_book as kb
        doc = (ROOT / "vcl" / "examples" / "example5_document.vcl").read_text(encoding="utf-8")
        result, out = kb.apply(doc, NEW, "x.xlsx", {}, TODAY)
        self.assertIsNone(out)
        self.assertFalse(result["ok"])


if __name__ == "__main__":
    unittest.main()
