"""VCL 0.7 part 3b (D113): merged cells.

A value written inside a merged range must go to its top-left cell; a row op on a sheet with merged cells stops; MERGE
merges a block (the same block again changes nothing, an overlap or a value Excel would drop stops the run); BLOCK
gives a block whose rows are known only at run time; MERGED(<Sheet>!<range>) checks a merge in a TEST. Each refusal
happens in the interpreter and in the compiled VBA alike (ForgeMergeCheck, ForgeNoMerges, ForgeMerge, run in the VBA
simulator).
"""
import datetime as dt
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(ROOT / "vcl"))

from validator import validate_text  # noqa: E402
from vba_check import check_text  # noqa: E402
from vcl_codegen import compile_text  # noqa: E402
from vcl_fixtures import Sheet  # noqa: E402
from vcl_interp import VclError, merge, run_tests  # noqa: E402

TODAY = dt.date(2026, 11, 20)
PRACTICE = (ROOT / "tests" / "data" / "merge_register.vcl").read_text(encoding="utf-8")

ONE_OP = """VCL 0.7
TASK
GOAL     Try one change on the practice WPS form, which has merged cells.
PROC     TryOne
KIND     Function
MODULE   modTry
ARGS     ws As Worksheet, go As Long
RET      Long
PRE      SAME(CELL($ws, 6, 1), "Welding Procedure Specification No. :")
MODIFIES $ws
OP       IF $go = 1
OP         {op}
OP       RETURN(1)
ERR      standard
FIXTURE  wps_03.xlsx template=wps_form
TEST     fixture=wps_03.xlsx call=TryOne(WPS, 1) assert={want} AND COUNTA(WPS!A1:I50) = PRE_COUNTA(WPS!A1:I50)
TEST     fixture=wps_03.xlsx call=TryOne(WPS, 1) runs=2 assert={want}
TEST     fixture=wps_03.xlsx call=TryOne(WPS, 0) assert=RESULT = 1 AND ERROR_COUNT = 0
"""
STOPS = "RESULT = -2 AND ERROR_COUNT = 1"
WORKS = "RESULT = 1 AND ERROR_COUNT = 0"


def errors(text):
    return [p.message for p in validate_text(text) if p.severity == "error"]


def proven_everywhere(test, text):
    test.assertEqual(errors(text), [])
    results = run_tests(text, TODAY)
    test.assertTrue(all(r.passed for r in results), [r.failed for r in results])
    outcome = check_text(text, "merged", TODAY)
    test.assertTrue(outcome.proven, outcome.differences or outcome.skipped)
    return results


class Writes(unittest.TestCase):
    def test_a_value_inside_a_merged_range_stops_the_run_and_its_top_left_cell_works(self):
        for op, want in (('SET_VALUE(RANGE($ws, "D6"), "x")', STOPS),        # C6:E6, not its top-left cell
                         ('SET_VALUE(RANGE($ws, "C6"), "DM-WPS-202")', WORKS),
                         ('WRITE_ROW($ws, 6, "a", "b", col=3)', STOPS),
                         ('CLEAR(RANGE($ws, "E6"), what=contents)', STOPS),
                         ('CLEAR(RANGE($ws, "E6"), what=formats)', WORKS),
                         ('SET_FORMULA(RANGE($ws, "D7"), "=1+1")', STOPS),
                         ('FORMAT(RANGE($ws, "D6"), fill=yellow)', WORKS)):
            with self.subTest(op=op):
                results = proven_everywhere(self, ONE_OP.format(op=op, want=want))
                if want == STOPS:
                    self.assertIn("merged", " ".join(m for r in results for _, _, m in r.log))

    def test_a_row_op_on_a_sheet_with_merged_cells_stops_the_run(self):
        for op in ("INSERT_ROW($ws, 40)", "LOOP_ROWS dir=up from=40 to=40 var=r\nOP           DELETE_ROW($ws, r)"):
            with self.subTest(op=op):
                proven_everywhere(self, ONE_OP.format(op=op, want=STOPS))


class Merge(unittest.TestCase):
    def test_the_practice_register_spec_merges_a_new_enquiry_in_every_stage(self):
        proven_everywhere(self, PRACTICE)
        vba = compile_text(PRACTICE, "merge_register")
        self.assertIn("ForgeMerge ws.Range(ws.Cells(r, c), ws.Cells((r + 1), c))", vba)

    def test_what_merge_refuses_in_every_stage(self):
        for op, want in (('MERGE(RANGE($ws, "D6:E6"))', STOPS),     # overlaps C6:E6
                         ('MERGE(RANGE($ws, "C6:E6"))', WORKS),     # the same block: nothing changes
                         ('MERGE(RANGE($ws, "A40:B40"))', WORKS),   # empty cells
                         ('MERGE(RANGE($ws, "C25:D25"))', STOPS),   # D25 holds "#"
                         ('MERGE(RANGE($ws, "A40"))', STOPS)):      # one cell
            with self.subTest(op=op):
                proven_everywhere(self, ONE_OP.format(op=op, want=want))

    def test_merge_in_the_model(self):
        sh = Sheet("S")
        sh.set(1, 1, "top")
        merge(sh, 1, 1, 2, 2)
        merge(sh, 1, 1, 2, 2)
        self.assertEqual(sh.merged, [(1, 1, 2, 2)])
        with self.assertRaises(VclError):
            merge(sh, 2, 2, 3, 3)

    def test_merge_needs_the_whole_sheet_and_vcl_07(self):
        narrow = ONE_OP.format(op='MERGE(RANGE($ws, "A40:B40"))', want=WORKS).replace(
            "MODIFIES $ws\n", 'MODIFIES $ws["Company Name :"]\n')
        self.assertTrue(errors(narrow))
        old = ONE_OP.format(op='MERGE(RANGE($ws, "A40:B40"))', want=WORKS).replace("VCL 0.7", "VCL 0.6", 1)
        self.assertTrue(any("VCL 0.7" in m for m in errors(old)), errors(old))

    def test_the_merged_assertion_is_vcl_07(self):
        old = PRACTICE.replace("VCL 0.7", "VCL 0.6", 1)
        self.assertTrue(any("MERGED is VCL 0.7" in m for m in errors(old)), errors(old))


if __name__ == "__main__":
    unittest.main()
