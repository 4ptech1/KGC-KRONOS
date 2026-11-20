"""VCL 0.7 part 3d (D114): two workbooks, `TEST fixture=<file> read=<other file> ...`.

The read workbook is a second FIXTURE file, passed to a Workbook argument; nothing can change it, so its formula
cells give the values Excel saved. Such a recipe is KRONOS Offline only: the validator says so, the compiler refuses
it, and the VBA simulator skips it (no macro may open a second workbook until March 2027, D36).
"""
import datetime as dt
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(ROOT / "vcl"))

from kronos_engine import prove  # noqa: E402
from validator import validate_text  # noqa: E402
from vba_check import check_text  # noqa: E402
from vcl_codegen import OFFLINE_ONLY, CompileError, compile_text  # noqa: E402
from vcl_fixtures import Formula, Sheet  # noqa: E402
from vcl_fmt import format_text  # noqa: E402
from vcl_interp import VclError, run_tests, value  # noqa: E402

TODAY = dt.date(2026, 11, 20)
PRACTICE = (ROOT / "tests" / "data" / "wps_from_wps.vcl").read_text(encoding="utf-8")


def problems(text, severity):
    return [p.message for p in validate_text(text) if p.severity == severity]


class TwoWorkbooks(unittest.TestCase):
    def test_the_practice_spec_passes_using_the_old_wps_saved_formula_values(self):
        self.assertEqual(problems(PRACTICE, "error"), [])
        self.assertEqual(format_text(PRACTICE), PRACTICE)   # read= keeps its place in the canonical text
        results = run_tests(PRACTICE, TODAY)
        self.assertTrue(all(r.passed for r in results), [r.failed for r in results])
        self.assertEqual(results[0].result, 3)   # the old WPS's T column: three formulas whose saved value is OK

    def test_it_is_kronos_offline_only(self):
        warnings = problems(PRACTICE, "warning")
        self.assertEqual(len([w for w in warnings if w.startswith("KRONOS Offline only")]), 1)
        with self.assertRaises(CompileError) as caught:
            compile_text(PRACTICE, "wps_from_wps")
        self.assertEqual(str(caught.exception), OFFLINE_ONLY)
        outcome = check_text(PRACTICE, "wps_from_wps", TODAY)
        self.assertFalse(outcome.proven)
        self.assertIn("KRONOS Offline only", outcome.skipped)
        proof = prove(PRACTICE, TODAY)
        self.assertFalse(proof.ok)
        self.assertIn("KRONOS Offline only", " ".join(f.message for f in proof.findings))

    def test_a_formula_with_no_saved_value_in_the_read_workbook_stops_the_run(self):
        sh = Sheet("WPS")
        sh.read_only = True
        sh.cells[(1, 1)] = Formula("=1+1", saved=2)
        sh.cells[(1, 2)] = Formula("=A1*2")
        self.assertEqual(value(sh, 1, 1), 2)
        with self.assertRaisesRegex(VclError, "no saved value"):
            value(sh, 1, 2)
        sh.read_only = False   # in a workbook the recipe may change, a formula is never read (R009)
        with self.assertRaisesRegex(VclError, "R009|can't calculate"):
            value(sh, 1, 1)

    def test_the_rules_of_read(self):
        old = PRACTICE.replace("VCL 0.7", "VCL 0.6", 1)
        self.assertTrue(any("read= is VCL 0.7" in m for m in problems(old, "error")))
        same = PRACTICE.replace("read=wps_old_01.xlsx call=FillWpsFromOld(wps_old_01.xlsx",
                                "read=wps_new_01.xlsx call=FillWpsFromOld(wps_new_01.xlsx")
        self.assertTrue(any("is the TEST's own workbook" in m for m in problems(same, "error")))
        missing = PRACTICE.replace("read=wps_old_01.xlsx", "read=wps_gone_01.xlsx")
        self.assertTrue(any("no FIXTURE line builds it" in m for m in problems(missing, "error")))
        written = PRACTICE.replace("MODIFIES $ws\n", "MODIFIES $old sheets\nMODIFIES $ws\n")
        self.assertTrue(any("read-only" in m for m in problems(written, "error")), problems(written, "error"))


if __name__ == "__main__":
    unittest.main()
