"""VCL 0.3 (docs/vep-0.3-file-ops.md), router Phase 1a: the version line, SAFE_NAME, PATH_JOIN
and FOLDER_EXISTS, and the path rules of the type checker (T4)."""
import glob
import os
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(ROOT / "vcl"))
import validator  # noqa: E402
from vcl_codegen import CompileError, compile_text  # noqa: E402
from vcl_interp import run_tests  # noqa: E402
from vcl_parse import parse  # noqa: E402
from vcl_types import check_text  # noqa: E402

# Job IDs typed by hand: two good, one padded, three unsafe (.., CON, A/B).
NAMES = r"""VCL 0.3
TASK
# ASSUME: Jobs headers (row 1) include "Job ID", "Safe Name" and "Path"
PROC     CheckJobNames
KIND     Function
MODULE   modJobs
ARGS     ws As Worksheet, root As String
RET      Long
VAR      cJob As Long
VAR      cName As Long
VAR      cPath As Long
VAR      n As Long
PRE      COL($ws, "Job ID") > 0; COL($ws, "Safe Name") > 0; COL($ws, "Path") > 0
MODIFIES $ws["Safe Name", "Path"]
OP       LET(cJob, COL($ws, "Job ID"))
OP       LET(cName, COL($ws, "Safe Name"))
OP       LET(cPath, COL($ws, "Path"))
OP       LET(n, 0)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cJob) var=r
OP         SET_VALUE(CELL($ws, r, cName), SAFE_NAME(CELL($ws, r, cJob)))
OP         SET_VALUE(CELL($ws, r, cPath), PATH_JOIN($root, SAFE_NAME(CELL($ws, r, cJob)), "Reports"))
OP         IF IS_EMPTY(CELL($ws, r, cName)) THEN LET(n, n + 1)
OP       RETURN(n)
ERR      standard
FIXTURE  jobs_01.xlsx sheet=Jobs cols=A:Job ID,B:Safe Name,C:Path rows=6 data=A:values("J-101"x2," J-102 "x1,".."x1,"CON"x1,"A/B"x1)
TEST     fixture=jobs_01.xlsx call=CheckJobNames(Jobs, "C:\Jobs") runs=2 assert=RESULT = 3 AND COUNTIF(Jobs![Path], "C:\Jobs\J-101\Reports") = 2 AND COUNTIF(Jobs![Safe Name], "J-102") = 1 AND ERROR_COUNT = 0
TEST     fixture=jobs_01.xlsx call=CheckJobNames(Jobs, "C:\Jobs") assert=COUNTA(Jobs![Job ID]) = PRE_COUNTA(Jobs![Job ID])
"""

HEAD = ["VCL 0.3", "TASK", "PROC     P", "KIND     Function", "MODULE   modP",
        "ARGS     ws As Worksheet, root As String", "RET      Long", "VAR      n As Long",
        "VAR      t As String", "VAR      flag As Boolean"]


def spec(*ops):
    return "\n".join(HEAD + ["OP       " + o for o in ops] + ["OP       RETURN(n)", "ERR      standard"])


def errors(text):
    return [p.message for p in validator.validate_text(text) if p.severity == "error"]


def type_problems(text):
    return [m for _, m in check_text(text)]


class VersionLine(unittest.TestCase):
    def test_both_parsers_read_it(self):
        self.assertEqual(parse(NAMES)[0]["version"], "VCL 0.3")
        self.assertEqual(validator.parse_spec(NAMES).version, "VCL 0.3")

    def test_every_vcl_02_spec_stays_valid_under_0_3(self):
        paths = sorted(glob.glob(os.path.join(ROOT, "specs", "*.vcl")))
        self.assertGreaterEqual(len(paths), 11)
        for path in paths:
            text = Path(path).read_text(encoding="utf-8")
            with self.subTest(spec=os.path.basename(path)):
                self.assertTrue(text.startswith("VCL 0.2\n"))
                newer = text.replace("VCL 0.2", "VCL 0.3", 1)
                self.assertEqual([(p.line, p.severity, p.message) for p in validator.validate_text(text)],
                                 [(p.line, p.severity, p.message) for p in validator.validate_text(newer)])
                self.assertEqual([(r.passed, r.result, r.state) for r in run_tests(text)],
                                 [(r.passed, r.result, r.state) for r in run_tests(newer)])

    def test_the_new_functions_need_the_0_3_line(self):
        older = NAMES.replace("VCL 0.3", "VCL 0.2", 1)
        for name in ("SAFE_NAME", "PATH_JOIN"):
            self.assertIn(f"{name} is VCL 0.3: start the spec with 'VCL 0.3' to use it", errors(older))
        self.assertIn("FOLDER_EXISTS is VCL 0.3: start the spec with 'VCL 0.3' to use it",
                      errors(spec("IF FOLDER_EXISTS($root) THEN LET(n, 1)").replace("VCL 0.3", "VCL 0.2", 1)))


class Validator(unittest.TestCase):
    def test_the_example_is_valid(self):
        self.assertEqual([(p.line, p.message) for p in validator.validate_text(NAMES)], [])

    def test_argument_counts(self):
        self.assertIn("PATH_JOIN takes 2 or more arguments, got 1",
                      errors(spec("LET(t, PATH_JOIN($root))")))
        self.assertIn("SAFE_NAME takes 1 argument, got 2", errors(spec('LET(t, SAFE_NAME("a", "b"))')))
        # The helper has no TESTs (rule 5); only the call's own problems matter here.
        self.assertEqual(errors(spec('LET(t, PATH_JOIN($root, "Jobs", "J-101", "r1.pdf"))')),
                         ["a spec needs at least two TEST lines (rule 5)"])


class PathTypes(unittest.TestCase):
    """T4: a path is a String $argument or PATH_JOIN($root, part, ...) with every part
    SAFE_NAME(...) or a quoted safe name, written in place."""

    def test_good_paths(self):
        for op in ('LET(t, PATH_JOIN($root, SAFE_NAME(CELL($ws, 2, 1))))',
                   'LET(t, PATH_JOIN($root, "Reports", SAFE_NAME(CELL($ws, 2, 1)), "r1.pdf"))',
                   'IF FOLDER_EXISTS($root) THEN LET(n, 1)',
                   'IF FOLDER_EXISTS(PATH_JOIN($root, "Reports")) THEN LET(n, 1)',
                   'SET_VALUE(CELL($ws, 2, 2), SAFE_NAME(CELL($ws, 2, 1)))'):
            with self.subTest(op=op):
                self.assertEqual(type_problems(spec(op)), [])

    def test_a_path_must_start_from_a_string_argument(self):
        for root in ('"C:\\Jobs"', "t", "$ws", "TRIM(t)"):
            with self.subTest(root=root):
                found = type_problems(spec(f'LET(t, PATH_JOIN({root}, "Reports"))'))
                self.assertTrue(any("must start from a String $argument" in m for m in found), found)

    def test_every_part_is_checked_where_it_is_used(self):
        for part, why in (('".."', "not a safe name"), ('"A\\B"', "not a safe name"), ('"CON"', "not a safe name"),
                          ("t", "must be SAFE_NAME(...)"), ("CELL($ws, 2, 1)", "must be SAFE_NAME(...)"),
                          ('TRIM(CELL($ws, 2, 1))', "must be SAFE_NAME(...)")):
            with self.subTest(part=part):
                found = type_problems(spec(f"LET(t, PATH_JOIN($root, {part}))"))
                self.assertTrue(any(why in m for m in found), found)

    def test_folder_exists_takes_a_path(self):
        for arg in ("t", '"C:\\Jobs"', "SAFE_NAME(t)"):
            with self.subTest(arg=arg):
                found = type_problems(spec(f"IF FOLDER_EXISTS({arg}) THEN LET(n, 1)"))
                self.assertTrue(any("is not a path" in m for m in found), found)

    def test_names_and_paths_are_read_as_text(self):
        for op in ('IF SAFE_NAME(t) = "" THEN LET(n, 1)', 'IF SAFE_NAME(t) <> t THEN LET(n, 1)',
                   'IF PATH_JOIN($root, "a") > t THEN LET(n, 1)', 'LET(t, FILE_NAME(PATH_JOIN($root, "a")))',
                   'LET(t, SAFE_NAME(t) & "x")', 'LET(t, IIF(flag, SAFE_NAME(t), "x"))',
                   'LET(n, LEN(PATH_JOIN($root, "a")))', 'LET(flag, FOLDER_EXISTS($root))'):
            with self.subTest(op=op):
                self.assertEqual(type_problems(spec(op)), [])
        self.assertIn("can't compare Name with Num", type_problems(spec('IF SAFE_NAME(t) = 1 THEN LET(n, 1)')))

    def test_names_and_paths_are_text_but_not_numbers(self):
        self.assertTrue(any("expected a number" in m or "not a number" in m
                            for m in type_problems(spec("LET(n, MOD(SAFE_NAME(t), 2))"))))
        self.assertTrue(any("expected text" in m for m in type_problems(spec("LET(t, SAFE_NAME($ws))"))))


class Interpreter(unittest.TestCase):
    def test_the_example_passes_its_tests(self):
        results = run_tests(NAMES)
        self.assertTrue(all(r.passed for r in results), [(r.line, r.failed, r.result) for r in results])
        self.assertEqual(results[0].result, 3)

    def test_the_compiler_waits_for_phase_3(self):
        with self.assertRaises(CompileError):
            compile_text(NAMES)


if __name__ == "__main__":
    unittest.main()
