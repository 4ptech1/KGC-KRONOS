import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
from vcl_types import check_text  # noqa: E402

HEAD = ["VCL 0.1", "TASK", "PROC     P", "KIND     Function", "MODULE   m",
        "ARGS     ws As Worksheet", "RET      Long", "VAR      n As Long", "VAR      t As String",
        "VAR      v As Variant", "VAR      d As Dictionary", "VAR      flag As Boolean"]


def spec(*ops):
    lines = HEAD + ["OP       " + o for o in ops] + ["OP       RETURN(n)", "ERR      standard"]
    return "\n".join(lines)


def loop(*body):
    return ["LOOP_ROWS dir=down from=2 to=LASTROW($ws, 1) var=r"] + ["  " + b for b in body]


class Catches(unittest.TestCase):
    def assertProblem(self, text, fragment):
        found = check_text(text)
        self.assertTrue(any(fragment in m for _, m in found), found)

    def test_guard_on_left_of_and_does_not_protect_right(self):
        self.assertProblem(spec(*loop("IF IS_NUMBER(CELL($ws, r, 2)) AND CELL($ws, r, 2) > 5 THEN LET(n, 1)")),
                           "compared with >")

    def test_raw_cell_into_long(self):
        self.assertProblem(spec(*loop("LET(n, CELL($ws, r, 2))")), "assigning raw cell value")

    def test_raw_cell_arithmetic(self):
        self.assertProblem(spec(*loop("SET_VALUE(CELL($ws, r, 3), CELL($ws, r, 2) * 2)")), "arithmetic")

    def test_write_removes_the_guard(self):
        self.assertProblem(spec(*loop("IF IS_NUMBER(CELL($ws, r, 2))",
                                      "  SET_VALUE(CELL($ws, r, 2), \"x\")",
                                      "  LET(n, CELL($ws, r, 2) + 1)")), "arithmetic")

    def test_sort_removes_all_guards_on_that_sheet(self):
        self.assertProblem(spec(*loop("IF IS_NUMBER(CELL($ws, r, 2))",
                                      "  SORT(USED_RANGE($ws), key=1, order=asc, header=yes)",
                                      "  LET(n, CELL($ws, r, 2) + 1)")), "arithmetic")

    def test_dict_sum_needs_guard(self):
        self.assertProblem(spec(*loop("DICT_ADD(d, TRIM(CELL($ws, r, 1)), CELL($ws, r, 2), mode=sum)")),
                           "mode=sum")

    def test_argument_without_dollar(self):
        self.assertProblem(spec("LET(n, LASTROW(ws, 1))"), "must be written $ws")

    def test_unguarded_date_comparison(self):
        self.assertProblem(spec(*loop("IF CELL($ws, r, 4) < TODAY THEN LET(n, n + 1)")), "compared with <")

    def test_cell_equals_number(self):
        self.assertProblem(spec(*loop("IF CELL($ws, r, 2) = 0 THEN LET(n, n + 1)")), "compared with a number")

    def test_flag_guard_is_not_provable(self):
        self.assertProblem(spec(*loop("LET(v, CELL($ws, r, 4))", "LET(flag, NOT IS_DATE(v))",
                                      "IF NOT flag", "  IF v < TODAY THEN LET(n, n + 1)")), "compared with <")

    def test_guard_from_outside_a_loop_does_not_carry_in(self):
        self.assertProblem(spec("IF IS_NUMBER(CELL($ws, 2, 2))", "  LET(n, CELL($ws, 2, 2))",
                                *loop("LET(n, CELL($ws, r, 2))")), "assigning raw cell value")


class Accepts(unittest.TestCase):
    def assertClean(self, text):
        self.assertEqual(check_text(text), [])

    def test_nested_guard(self):
        self.assertClean(spec(*loop("IF IS_NUMBER(CELL($ws, r, 2))", "  IF CELL($ws, r, 2) > 5 THEN LET(n, n + 1)")))

    def test_one_line_if_guard(self):
        self.assertClean(spec(*loop("IF IS_NUMBER(CELL($ws, r, 2)) THEN SET_VALUE(CELL($ws, r, 3), CELL($ws, r, 2) * 2)")))

    def test_else_branch_guard(self):
        self.assertClean(spec(*loop("IF NOT IS_NUMBER(CELL($ws, r, 2))", "  LET(n, n + 1)", "ELSE",
                                    "  SET_VALUE(CELL($ws, r, 3), CELL($ws, r, 2) * 2)")))

    def test_exit_if_guard(self):
        self.assertClean(spec("EXIT_IF NOT IS_DATE(CELL($ws, 2, 4))", "IF CELL($ws, 2, 4) < TODAY THEN LET(n, 1)"))

    def test_write_to_a_different_cell_keeps_the_guard(self):
        self.assertClean(spec(*loop("IF IS_NUMBER(CELL($ws, r, 2))", "  SET_VALUE(CELL($ws, r, 3), 0)",
                                    "  LET(n, CELL($ws, r, 2) + 1)")))

    def test_text_comparison_needs_no_guard(self):
        self.assertClean(spec(*loop('IF UPPER(TRIM(CELL($ws, r, 2))) = "OPEN" THEN LET(n, n + 1)',
                                    'IF CELL($ws, r, 2) = "Open" THEN LET(n, n + 1)')))

    def test_date_arithmetic(self):
        self.assertClean(spec(*loop("IF IS_DATE(CELL($ws, r, 4))",
                                    "  IF CELL($ws, r, 4) <= TODAY + 15 THEN LET(n, TODAY - CELL($ws, r, 4))")))

    def test_real_specs_except_task_06(self):
        for p in sorted((ROOT / "specs").glob("*.vcl")):
            found = check_text(p.read_text())
            if p.name == "task_06.vcl":
                self.assertEqual(len(found), 1)
            else:
                self.assertEqual(found, [], p.name)


if __name__ == "__main__":
    unittest.main()
