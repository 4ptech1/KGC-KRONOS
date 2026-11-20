"""VCL 0.7 part 3c (D111): ROUND, ROUNDUP, MIN and MAX as value functions.

Their meaning is Excel's (vcl/semantics.py S27, S28; conformance cases in tests/conformance_cases.py). The validator
gates them to VCL 0.7 and wants ROUND's digits written in place; the type checker wants numbers guarded (rule 15); the
interpreter, the compiler (ForgeLib's ForgeRound, ForgeRoundUp, ForgeMin, ForgeMax) and the VBA simulator agree.
"""
import datetime as dt
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(ROOT / "vcl"))

import semantics as S  # noqa: E402
import vba_sim as V  # noqa: E402
from validator import validate_text  # noqa: E402
from vba_check import check_text  # noqa: E402
from vcl_codegen import compile_text  # noqa: E402
from vcl_interp import run_tests  # noqa: E402

TODAY = dt.date(2026, 11, 20)
PRACTICE = (ROOT / "tests" / "data" / "rule_functions.vcl").read_text(encoding="utf-8")

SMALL = """VCL 0.7
TASK
GOAL     Write the smaller and the larger of two numbers, and a rounded value, to the first row.
PROC     WriteBounds
KIND     Sub
MODULE   modBounds
ARGS     ws As Worksheet, a As Double, b As Double
PRE      COL($ws, "Low") = 1; COL($ws, "High") = 2; COL($ws, "Round") = 3
MODIFIES $ws["Low", "High", "Round"]
OP       SET_VALUE(CELL($ws, 2, 1), MIN($a, $b))
OP       SET_VALUE(CELL($ws, 2, 2), MAX($a, $b))
OP       SET_VALUE(CELL($ws, 2, 3), {round})
ERR      standard
FIXTURE  bounds_01.xlsx sheet=Out cols=A:Low,B:High,C:Round rows=1 data=A:values("x"x1)
TEST     fixture=bounds_01.xlsx call=WriteBounds(Out, 2.5, -4) runs=2 assert=VALUE(Out!A2) = -4 AND VALUE(Out!B2) = 2.5 AND ERROR_COUNT = 0
TEST     fixture=bounds_01.xlsx call=WriteBounds(Out, 7, 7) assert=VALUE(Out!A2) = 7 AND VALUE(Out!B2) = 7 AND COUNTA(Out![Low]) = PRE_COUNTA(Out![Low])
"""


def codes(text):
    return [p.code for p in validate_text(text) if p.severity == "error"]


class TheValidator(unittest.TestCase):
    def test_the_practice_spec_is_valid(self):
        self.assertEqual(codes(PRACTICE), [])
        self.assertEqual(codes(SMALL.format(round="ROUND($a * 1.01972, 2)")), [])

    def test_they_are_vcl_07(self):
        found = validate_text(PRACTICE.replace("VCL 0.7", "VCL 0.6", 1))
        self.assertTrue(any("is VCL 0.7" in p.message for p in found if p.severity == "error"), found)

    def test_the_digits_are_a_whole_number_written_in_place(self):
        for bad in ("ROUND($a, $b)", "ROUND($a, 2.5)", "ROUND($a, 16)", "ROUNDUP($a, -16)"):
            with self.subTest(expr=bad):
                found = [p.message for p in validate_text(SMALL.format(round=bad)) if p.severity == "error"]
                self.assertTrue(any("whole number from -15 to 15" in m for m in found), found)
        self.assertEqual(codes(SMALL.format(round="ROUNDUP($a, -3)")), [])

    def test_a_cell_is_guarded_before_it_is_rounded(self):
        bare = PRACTICE.replace("OP         IF IS_NUMBER(CELL($ws, r, cBar))\n", "OP         IF TRUE\n")
        self.assertTrue(codes(bare))


class TheMeaning(unittest.TestCase):
    def test_the_practice_spec_passes_and_the_vba_agrees(self):
        results = run_tests(PRACTICE, TODAY)
        self.assertTrue(all(r.passed for r in results), [r.failed for r in results])
        vba = compile_text(PRACTICE, "rule_functions")
        self.assertIn("ForgeRound(", vba)
        self.assertIn("ForgeRoundUp(", vba)
        self.assertIn("ForgeMax(", vba)
        outcome = check_text(PRACTICE, "rule_functions", TODAY)
        self.assertTrue(outcome.proven, outcome.differences or outcome.skipped)

    def test_min_and_max_in_every_stage(self):
        text = SMALL.format(round="ROUND($a * 1.01972, 2)")
        results = run_tests(text, TODAY)
        self.assertTrue(all(r.passed for r in results), [r.failed for r in results])
        self.assertIn("ForgeMin(", compile_text(text, "bounds"))
        outcome = check_text(text, "bounds", TODAY)
        self.assertTrue(outcome.proven, outcome.differences or outcome.skipped)

    def test_the_simulators_worksheet_functions_are_excels(self):
        sim = V.Simulator([V.Module("Public Function T(ByVal x As Double, ByVal d As Long) As Double\n"
                                    "    T = Application.WorksheetFunction.RoundUp(x, d) + "
                                    "Application.WorksheetFunction.Round(x, d)\nEnd Function\n")], today=TODAY)
        self.assertEqual(sim.call("T", 288364.0, -3), S.round_up(288364, -3) + S.round_half(288364, -3))
        self.assertEqual(sim.call("T", 2.675, 2), 2.68 + 2.68)


if __name__ == "__main__":
    unittest.main()
