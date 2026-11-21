"""KRONOS Offline version 6, part 6b (D119): a changed copy that keeps everything.

When a recipe moves no row and no sheet, the copy is the uploaded file with only the changed cells written into its
XML (compiler/kronos_patch.py): every other part stays byte for byte, so the WPS look-alike keeps its pictures, print
area, page break and merged cells. Fills and number formats get a new cell style beside the cell's own; a changed
formula drops the calculation chain; Excel recalculates on opening. A row that moves still gets the full copy.
(The round trip of every recipe's TESTs in tests/test_kronos_book.py runs through both writers.)
"""
import datetime as dt
import io
import re
import sys
import unittest
import zipfile
import xml.etree.ElementTree as ET
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
WPS = (ROOT / "data" / "templates" / "wps_form.xlsx").read_bytes()
REGISTER = (ROOT / "data" / "templates" / "enquiry_register.xlsx").read_bytes()
M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
WRITTEN = ["H2", "H3", "H35", "H36"]   # the recipe writes the revision and date in both halves of the form
REVISION = (ROOT / "tests" / "data" / "template_wps.vcl").read_text(encoding="utf-8")


def parts(data):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return {n: z.read(n) for n in z.namelist()}


def task(ops, args="ws As Worksheet", modifies="MODIFIES $ws", sheet="WPS", fixture="wps_09.xlsx template=wps_form"):
    """A small TASK on a template, with its two TESTs."""
    return (f"VCL 0.7\nTASK\nGOAL     Change a few cells of the practice form for the keep-everything tests.\n"
            f"PROC     ChangeForm\nKIND     Sub\nMODULE   modForm\nARGS     {args}\n{modifies}\n"
            + "".join(f"OP       {op}\n" for op in ops) +
            f"ERR      standard\nFIXTURE  {fixture}\n"
            f"TEST     fixture={fixture.split()[0]} call=ChangeForm({sheet}) assert=ERROR_COUNT = 0\n"
            f"TEST     fixture={fixture.split()[0]} call=ChangeForm({sheet}) runs=2 assert=COUNTA({sheet}!A1:A9) = "
            f"PRE_COUNTA({sheet}!A1:A9)\n")


@unittest.skipIf(openpyxl is None, "openpyxl is not installed")
class KeepsEverything(unittest.TestCase):
    def apply(self, text, data, sheet="WPS", name="form.xlsx"):
        import kronos_book as kb
        result, out = kb.apply(text, data, name, {"ws": sheet}, TODAY)
        self.assertTrue(result["ok"], result["message"])
        return result, out

    def test_the_wps_comes_back_byte_for_byte_but_for_its_changed_cells(self):
        import kronos_book as kb
        result, out = kb.apply(REVISION, WPS, "wps_form.xlsx", {"ws": "WPS", "rev": "R1", "dateText": "20/11/2026"},
                               TODAY)
        self.assertTrue(result["ok"] and result["kept_everything"], result["message"])
        self.assertEqual(result["warnings"], [])
        before, after = parts(WPS), parts(out)
        self.assertEqual(set(before), set(after))
        changed = sorted(n for n in before if before[n] != after[n])
        self.assertEqual(changed, ["xl/workbook.xml", "xl/worksheets/sheet1.xml"])
        self.assertIn(b'fullCalcOnLoad="1"', after["xl/workbook.xml"])
        self.assertTrue(any(n.startswith("xl/media/") for n in after))   # the logo and the sketch
        ws = openpyxl.load_workbook(io.BytesIO(out))["WPS"]
        self.assertEqual((ws["H3"].value, ws["H36"].value, ws["H2"].value),
                         ("WPS Revision : R1", "WPS Revision : R1", "DATE:- 20/11/2026"))
        self.assertEqual(ws.print_area, "'WPS'!$A$1:$I$50")
        self.assertEqual(len(ws.merged_cells.ranges), 97)
        self.assertEqual(ws["R29"].value, "=$N$25*N29*L29*60/(Q29*1000)")   # a formula nobody changed stays
        # the sheet's XML differs only in the cells written: every other element, row and cell is as it was
        old, new = ET.fromstring(before["xl/worksheets/sheet1.xml"]), ET.fromstring(after["xl/worksheets/sheet1.xml"])
        self.assertEqual(after["xl/worksheets/sheet1.xml"][:200].count(b"ns0"), 0)   # its namespace is the default one
        self.assertEqual([e.tag for e in old], [e.tag for e in new])
        for a, b in zip(old, new):
            if a.tag != M + "sheetData":
                self.assertEqual(ET.tostring(a), ET.tostring(b))
        cells = lambda root: {c.get("r"): ET.tostring(c) for c in root.iter(M + "c")}  # noqa: E731
        rows = lambda root: [(r.get("r"), sorted(r.attrib.items())) for r in root.iter(M + "row")]  # noqa: E731
        self.assertEqual(rows(old), rows(new))
        a, b = cells(old), cells(new)
        self.assertEqual(sorted(ref for ref in set(a) | set(b) if a.get(ref) != b.get(ref)), WRITTEN)

    def test_a_fill_and_a_number_format_get_their_own_cell_style(self):
        text = task(['FORMAT(RANGE($ws, "C25"), fill=yellow, numfmt="0.0")', 'SET_VALUE(RANGE($ws, "C6"), 7.25)'])
        result, out = self.apply(text, WPS)
        self.assertTrue(result["kept_everything"])
        ws = openpyxl.load_workbook(io.BytesIO(out))["WPS"]
        self.assertEqual((ws["C25"].fill.fgColor.rgb, ws["C25"].number_format, ws["C25"].value), ("FFFFFF00", "0.0", 12))
        self.assertEqual(ws["C25"].font.b, ws["D25"].font.b)   # the rest of its look is its own
        self.assertEqual(ws["C6"].value, 7.25)
        self.assertNotEqual(ws["C24"].fill.fgColor.rgb, "FFFFFF00")   # its neighbours keep theirs

    def test_cells_the_form_did_not_have_are_added_in_order(self):
        text = task(['SET_VALUE(RANGE($ws, "K60"), "note")', 'SET_VALUE(RANGE($ws, "B62"), TRUE)'])
        result, out = self.apply(text, WPS)
        ws = openpyxl.load_workbook(io.BytesIO(out))["WPS"]
        self.assertEqual((ws["K60"].value, ws["B62"].value), ("note", True))

    def test_a_changed_formula_drops_the_calculation_chain(self):
        src = parts(WPS)
        src["xl/calcChain.xml"] = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<calcChain xmlns="http://'
                                   b'schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="R29" i="1"/></calcChain>')
        src["xl/_rels/workbook.xml.rels"] = src["xl/_rels/workbook.xml.rels"].replace(
            b"</Relationships>", b'<Relationship Id="rIdChain" Type="http://schemas.openxmlformats.org/officeDocument/'
            b'2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>')
        src["[Content_Types].xml"] = src["[Content_Types].xml"].replace(
            b"</Types>", b'<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-'
            b'officedocument.spreadsheetml.calcChain+xml"/></Types>')
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as z:
            for n, b in src.items():
                z.writestr(n, b)
        result, out = self.apply(task(['SET_VALUE(RANGE($ws, "R29"), 0.5)']), buf.getvalue())
        after = parts(out)
        self.assertTrue(result["kept_everything"])
        self.assertNotIn("xl/calcChain.xml", after)
        self.assertNotIn(b"calcChain", after["xl/_rels/workbook.xml.rels"] + after["[Content_Types].xml"])
        self.assertEqual(openpyxl.load_workbook(io.BytesIO(out))["WPS"]["R29"].value, 0.5)

    def test_rows_that_move_still_get_the_full_copy(self):
        from test_kronos_book import EXAMPLE, orders, save
        import kronos_book as kb
        result, out = kb.apply(EXAMPLE, save(orders()), "orders.xlsx", {"ws": "Orders"}, TODAY)
        self.assertTrue(result["ok"], result["message"])
        self.assertFalse(result["kept_everything"])
        self.assertEqual(openpyxl.load_workbook(io.BytesIO(out))["Orders"].max_row, 11)

    def test_a_row_op_on_a_form_with_merged_cells_is_refused(self):
        import kronos_book as kb
        text = task(["LOOP_ROWS dir=up from=7 to=7 var=r", "  DELETE_ROW($ws, r)"])
        result, out = kb.apply(text, WPS, "form.xlsx", {"ws": "WPS"}, TODAY)
        self.assertIsNone(out)   # its own TESTs fail too: the macro refuses it (D113)
        self.assertFalse(result["ok"])

if __name__ == "__main__":
    unittest.main()
