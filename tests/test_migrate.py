"""compiler/vcl_migrate.py: VCL 0.1 to 0.2, only rewrites that can't change what a spec does."""
import datetime as dt
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
from vcl_migrate import MigrateError, migrate, prove, rewrite_same  # noqa: E402

V01 = ROOT / "tests" / "data" / "specs_v01"
TODAY = dt.date(2026, 10, 17)


def task(*ops, head=("VAR      n As Long",), mods=()):
    return "\n".join(["VCL 0.1", "TASK", "PROC     P", "KIND     Function", "MODULE   modX",
                      "ARGS     ws As Worksheet, out As Worksheet", "RET      Long", *head,
                      'PRE      COL($ws, "Status") > 0', *mods]
                     + ["OP       " + o for o in ops] + ["OP       RETURN(n)", "ERR      standard"]) + "\n"


def ops_of(text):
    return [line[9:] for line in text.splitlines() if line.startswith("OP ")]


class TheTenSpecs(unittest.TestCase):
    def test_migrating_the_vcl_01_specs_gives_the_committed_specs(self):
        for old in sorted(V01.glob("task_*.vcl")):
            with self.subTest(spec=old.name):
                new, _ = migrate(old.read_text(encoding="utf-8"))
                self.assertEqual(new, (ROOT / "specs" / old.name).read_text(encoding="utf-8"))

    def test_migrating_example_1_gives_the_prompts_example(self):
        # The tool migrates to 0.2; the prompt's Example 1 then moved to 0.3, 0.4 and 0.5 by its
        # first line only, which gives the same results (E8, D33; E12, D46; E13, D58).
        new, _ = migrate((V01 / "example1_rev3.vcl").read_text(encoding="utf-8"))
        example = (ROOT / "vcl" / "examples" / "example1_rev3.vcl").read_text(encoding="utf-8")
        self.assertTrue(new.startswith("VCL 0.2\n") and example.startswith("VCL 0.5\n"))
        self.assertEqual(new.replace("VCL 0.2\n", "VCL 0.5\n", 1), example)

    def test_the_rewrites_give_the_same_results(self):
        # The proof without its slow mutation part (vcl_migrate.py --prove runs that too).
        for old in sorted(V01.glob("*.vcl")):
            with self.subTest(spec=old.name):
                text = old.read_text(encoding="utf-8")
                problems, _ = prove(text, migrate(text)[0], TODAY, mutation=False)
                self.assertEqual(problems, [])

    def test_a_vcl_02_spec_is_left_alone(self):
        with self.assertRaises(MigrateError):
            migrate((ROOT / "specs" / "task_01.vcl").read_text(encoding="utf-8"))


class Same(unittest.TestCase):
    def test_rewrites(self):
        cases = {
            'IF UPPER(TRIM(CELL($ws, r, c))) = "OPEN" THEN LET(n, 1)': 'IF SAME(CELL($ws, r, c), "OPEN") THEN LET(n, 1)',
            'IF UPPER(TRIM(a)) <> "X" OR UPPER(TRIM(b)) <> "Y"': 'IF NOT SAME(a, "X") OR NOT SAME(b, "Y")',
            'IF NOT UPPER(TRIM(a)) = "X"': 'IF NOT SAME(a, "X")',
        }
        for before, after in cases.items():
            self.assertEqual(rewrite_same(before)[0], after)

    def test_leaves_what_it_cant_prove(self):
        for value in ['IF UPPER(TRIM(a)) = "Open" THEN LET(n, 1)',     # not upper case
                      'IF UPPER(TRIM(a)) = " OPEN" THEN LET(n, 1)',    # not trimmed
                      'IF UPPER(TRIM(a)) = "STRAßE" THEN LET(n, 1)',   # not ASCII
                      'IF UPPER(TRIM(a)) = "" THEN LET(n, 1)',         # empty
                      'IF UPPER(TRIM(a)) & "X" = "AX" THEN LET(n, 1)',  # part of a longer text
                      'IF t & UPPER(TRIM(a)) = "AX" THEN LET(n, 1)',
                      'IF UPPER(TRIM(a)) = "A" & t THEN LET(n, 1)',
                      'IF UPPER(TRIM(a)) <= "A" THEN LET(n, 1)',
                      'LET(t, UPPER(TRIM(a)))']:
            self.assertEqual(rewrite_same(value), (value, 0), value)


class SkipIf(unittest.TestCase):
    LOOP = "LOOP_ROWS dir=down from=2 to=LASTROW($ws, 1) var=r"

    def test_whole_body_guard_becomes_skip_if(self):
        new, _ = migrate(task(self.LOOP, "  IF NOT IS_EMPTY(CELL($ws, r, 1))", "    LET(n, n + 1)",
                              "    IF n > 3", "      LET(n, 0)", "    ELSE", "      LET(n, 1)"))
        self.assertEqual(ops_of(new)[:6], [self.LOOP, "  SKIP_IF IS_EMPTY(CELL($ws, r, 1))", "  LET(n, n + 1)",
                                           "  IF n > 3", "    LET(n, 0)", "  ELSE"])

    def test_other_bodies_stay(self):
        for body in (["  IF NOT IS_EMPTY(CELL($ws, r, 1))", "    LET(n, n + 1)", "  LET(n, n + 2)"],   # more steps
                     ["  IF NOT IS_EMPTY(CELL($ws, r, 1))", "    LET(n, n + 1)", "  ELSE", "    LET(n, 0)"],
                     ["  IF NOT IS_EMPTY(CELL($ws, r, 1)) AND n > 0", "    LET(n, n + 1)"],
                     ["  IF NOT IS_EMPTY(CELL($ws, r, 1)) THEN LET(n, n + 1)"]):
            with self.subTest(body=body):
                new, _ = migrate(task(self.LOOP, *body))
                self.assertEqual(ops_of(new)[1:len(body) + 1], body)


class WriteRow(unittest.TestCase):
    def test_runs_become_write_row(self):
        new, _ = migrate(task('SET_VALUE(CELL($out, 1, 2), "A")', 'SET_VALUE(CELL($out, 1, 3), n)',
                              'SET_VALUE(CELL($out, 1, 4), CELL($ws, 2, 1))', 'SET_VALUE(CELL($out, 1, 6), "B")'))
        self.assertEqual(ops_of(new)[:2], ['WRITE_ROW($out, 1, "A", n, CELL($ws, 2, 1), col=2)',
                                           'SET_VALUE(CELL($out, 1, 6), "B")'])

    def test_only_when_evaluating_first_changes_nothing(self):
        for second in ['SET_VALUE(CELL($out, 1, 2), CELL($out, 1, 1))',   # reads what the run writes
                       'SET_VALUE(CELL($out, 1, 2), n / 2)',              # could raise an error
                       'SET_VALUE(CELL($out, 2, 2), "B")',                # another row
                       'SET_VALUE(CELL($ws, 1, 2), "B")']:                # another sheet
            with self.subTest(second=second):
                new, _ = migrate(task('SET_VALUE(CELL($out, 1, 1), "A")', second))
                self.assertEqual(ops_of(new)[:2], ['SET_VALUE(CELL($out, 1, 1), "A")', second])


class Modifies(unittest.TestCase):
    def test_lines_from_the_writes(self):
        head = ("VAR      n As Long", "VAR      c As Long")
        new, _ = migrate(task('LET(c, COL($ws, "Status"))', "SET_VALUE(CELL($ws, 2, c), 1)",
                              "FORMAT(ROW_RANGE($ws, 2, 1, 3), fill=#FF0000)", "CLEAR(USED_RANGE($out), what=all)",
                              head=head))
        self.assertIn('MODIFIES $ws["Status"]\nMODIFIES $ws fills\nMODIFIES $out\nOP ', new)

    def test_a_column_it_cant_name_means_the_whole_sheet(self):
        new, _ = migrate(task("SET_VALUE(CELL($ws, 2, n + 1), 1)"))
        self.assertIn("MODIFIES $ws\nOP ", new)

    def test_a_write_it_cant_follow_is_refused(self):
        text = task("SET_VALUE(target, 1)", head=("VAR      n As Long", "VAR      target As Range"))
        with self.assertRaises(MigrateError):
            migrate(text)


if __name__ == "__main__":
    unittest.main()
