"""VCL 0.2 value functions (docs/vep-0.2-core.md, PR A): SAME, CLEAN_TEXT, IS_ERROR,
PARSE_DATE, IIF and MOD, through every stage: the version line, the validator, the type
checker, the interpreter, the compiler and ForgeLib."""
import glob
import os
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
sys.path.insert(0, str(ROOT / "vcl"))
sys.path.insert(0, str(ROOT / "tests"))
import semantics  # noqa: E402
import validator  # noqa: E402
from vba_lint import lint  # noqa: E402
from vcl_codegen import assemble, compile_text  # noqa: E402
from vcl_interp import run_tests  # noqa: E402
from vcl_parse import parse  # noqa: E402
from vcl_types import check_text  # noqa: E402

TIDY = """VCL 0.2
TASK
# ASSUME: Data headers (row 1) include "Status", "Due" and "Qty"; columns D to F are free
PROC     TidyRows
KIND     Function
MODULE   modTidy
ARGS     ws As Worksheet
RET      Long
VAR      cStatus As Long
VAR      cDue As Long
VAR      cQty As Long
VAR      n As Long
VAR      d As Variant
PRE      COL($ws, "Status") > 0; COL($ws, "Due") > 0; COL($ws, "Qty") > 0
OP       LET(cStatus, COL($ws, "Status"))
OP       LET(cDue, COL($ws, "Due"))
OP       LET(cQty, COL($ws, "Qty"))
OP       LET(n, 0)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cStatus) var=r
OP         IF SAME(CELL($ws, r, cStatus), "Ready") THEN LET(n, n + 1)
OP         SET_VALUE(CELL($ws, r, 4), CLEAN_TEXT(CELL($ws, r, cStatus)))
OP         LET(d, PARSE_DATE(CELL($ws, r, cDue), "dd/mm/yyyy"))
OP         IF IS_DATE(d) THEN SET_VALUE(CELL($ws, r, 5), d)
OP         IF IS_NUMBER(CELL($ws, r, cQty))
OP           SET_VALUE(CELL($ws, r, 6), IIF(MOD(CELL($ws, r, cQty), 2) = 0, "Even", "Odd"))
OP         IF IS_ERROR(CELL($ws, r, cQty)) THEN LET(n, n + 100)
OP       RETURN(n)
ERR      standard
FIXTURE  tidy_01.xlsx sheet=Data cols=A:Status,B:Due,C:Qty rows=6 data=A:values(" ready "x2,"READY"x1,"Closed"x3);B:values("15/10/2026"x2,"31/02/2026"x1,"2026-10-15"x3);C:int(1,9)
TEST     fixture=tidy_01.xlsx call=TidyRows(Data) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Data!D:D, "ready") = 3 AND COUNTA(Data!E:E) = 2 AND COUNTIF(Data!F:F, "Even") + COUNTIF(Data!F:F, "Odd") = 6
TEST     fixture=tidy_01.xlsx call=TidyRows(Data) assert=COUNTA(Data![Status]) = PRE_COUNTA(Data![Status]) AND COUNTA(Data![Due]) = PRE_COUNTA(Data![Due])
"""

HEAD = ["VCL 0.2", "TASK", "PROC     P", "KIND     Function", "MODULE   m",
        "ARGS     ws As Worksheet", "RET      Long", "VAR      n As Long", "VAR      t As String",
        "VAR      v As Variant", "VAR      flag As Boolean"]


def spec(*ops):
    return "\n".join(HEAD + ["OP       " + o for o in ops] + ["OP       RETURN(n)", "ERR      standard"])


def loop(*body):
    return ["LOOP_ROWS dir=down from=2 to=LASTROW($ws, 1) var=r"] + ["  " + b for b in body]


def errors(text):
    return [p.message for p in validator.validate_text(text) if p.severity == "error"]


class VersionLine(unittest.TestCase):
    def test_both_parsers_read_the_version(self):
        self.assertEqual(parse(TIDY)[0]["version"], "VCL 0.2")
        self.assertEqual(validator.parse_spec(TIDY).version, "VCL 0.2")
        old = TIDY.replace("VCL 0.2", "VCL 0.1", 1)
        self.assertEqual(parse(old)[0]["version"], "VCL 0.1")
        self.assertEqual(validator.parse_spec(old).version, "VCL 0.1")

    def test_unknown_version_is_an_error(self):
        found = errors(TIDY.replace("VCL 0.2", "VCL 0.3", 1))
        self.assertIn("the first line must be 'VCL 0.1' or 'VCL 0.2'", found)

    def test_every_vcl_01_spec_stays_valid_under_0_2(self):
        for path in sorted(glob.glob(os.path.join(ROOT, "specs", "*.vcl"))):
            text = Path(path).read_text(encoding="utf-8")
            with self.subTest(spec=os.path.basename(path)):
                self.assertTrue(text.startswith("VCL 0.1\n"))
                before = [(p.line, p.severity, p.message) for p in validator.validate_text(text)]
                after = [(p.line, p.severity, p.message)
                         for p in validator.validate_text(text.replace("VCL 0.1", "VCL 0.2", 1))]
                self.assertEqual(before, after)
                self.assertEqual([r.passed for r in run_tests(text)],
                                 [r.passed for r in run_tests(text.replace("VCL 0.1", "VCL 0.2", 1))])

    def test_a_build_keeps_one_version(self):
        from test_parser import MINIMAL_BUILD
        self.assertEqual(errors(MINIMAL_BUILD), [])
        mixed = MINIMAL_BUILD.replace("VCL 0.1", "VCL 0.2", 1)
        self.assertTrue(any("starts with 'VCL 0.2' and 'TASK'" in m for m in errors(mixed)), errors(mixed))
        whole = MINIMAL_BUILD.replace("VCL 0.1", "VCL 0.2")
        self.assertEqual(errors(whole), [])


class Validator(unittest.TestCase):
    def test_the_example_is_valid(self):
        self.assertEqual(errors(TIDY), [])

    def test_new_functions_need_the_0_2_version_line(self):
        found = errors(TIDY.replace("VCL 0.2", "VCL 0.1", 1))
        for name in ("SAME", "CLEAN_TEXT", "PARSE_DATE", "IIF", "MOD", "IS_ERROR"):
            self.assertIn(f"{name} is VCL 0.2: start the spec with 'VCL 0.2' to use it", found)

    def test_parse_date_needs_a_listed_format(self):
        found = errors(TIDY.replace('"dd/mm/yyyy"', '"d/m/yy"'))
        self.assertIn('PARSE_DATE format "d/m/yy" is not one of "dd/mm/yyyy", "dd-mm-yyyy", '
                      '"dd.mm.yyyy", "mm/dd/yyyy", "yyyy-mm-dd"', found)
        found = errors(TIDY.replace('"dd/mm/yyyy"', 'cDue'))
        self.assertIn('PARSE_DATE needs a "quoted" text as its second argument', found)
        for fmt in semantics.DATE_FORMATS:
            self.assertEqual(errors(TIDY.replace('"dd/mm/yyyy"', f'"{fmt}"')), [], fmt)

    def test_the_formats_are_the_reference_ones(self):
        self.assertEqual(validator.DATE_FORMATS, semantics.DATE_FORMATS)

    def test_argument_counts(self):
        self.assertIn("SAME takes 2 arguments, got 1",
                      errors(TIDY.replace('SAME(CELL($ws, r, cStatus), "Ready")', "SAME(CELL($ws, r, cStatus))")))
        self.assertIn("IIF takes 3 arguments, got 2", errors(TIDY.replace(', "Even", "Odd")', ', "Even")')))

    def test_function_names_need_brackets(self):
        self.assertIn("MOD needs its brackets: MOD(...)", errors(TIDY.replace("LET(n, 0)", "LET(n, MOD)")))


class Types(unittest.TestCase):
    def assertProblem(self, text, fragment):
        found = check_text(text)
        self.assertTrue(any(fragment in m for _, m in found), found)

    def assertClean(self, text):
        self.assertEqual(check_text(text), [])

    def test_iif_condition_does_not_guard_its_values(self):
        # Both values are always evaluated, so a text cell would still hit the * 2.
        self.assertProblem(spec(*loop("SET_VALUE(CELL($ws, r, 3), IIF(IS_NUMBER(CELL($ws, r, 2)), CELL($ws, r, 2) * 2, 0))")),
                           "arithmetic")

    def test_iif_type(self):
        self.assertClean(spec(*loop('LET(t, IIF(IS_EMPTY(CELL($ws, r, 2)), "Missing", "OK"))')))
        self.assertClean(spec("LET(n, IIF(flag, 1, 2))"))
        self.assertProblem(spec('LET(n, IIF(flag, 1, "x"))'), "assigning raw cell value")
        self.assertProblem(spec("LET(n, IIF(CELL($ws, 2, 2), 1, 2))"), "expected Bool")
        self.assertProblem(spec("LET(n, IIF(flag, 1, $ws))"), "expected a value")

    def test_parse_date_needs_an_is_date_guard(self):
        self.assertProblem(spec(*loop('LET(v, PARSE_DATE(CELL($ws, r, 2), "dd/mm/yyyy"))',
                                      "IF v < TODAY THEN LET(n, n + 1)")), "compared with <")
        self.assertClean(spec(*loop('LET(v, PARSE_DATE(CELL($ws, r, 2), "dd/mm/yyyy"))',
                                    "IF IS_DATE(v)", "  IF v < TODAY THEN LET(n, n + 1)")))

    def test_mod_needs_numbers(self):
        self.assertProblem(spec(*loop("LET(n, MOD(CELL($ws, r, 2), 2))")), "guard it with IS_NUMBER first")
        self.assertProblem(spec("LET(n, MOD(TODAY, 7))"), "expected a number")
        self.assertProblem(spec('LET(n, MOD(t, 7))'), "expected a number")
        self.assertClean(spec(*loop("IF IS_NUMBER(CELL($ws, r, 2)) THEN LET(n, MOD(CELL($ws, r, 2), 2))")))
        self.assertClean(spec("LET(n, MOD(n + 7, 3))"))

    def test_text_functions_need_no_guard(self):
        self.assertClean(spec(*loop('IF SAME(CELL($ws, r, 2), "Ready") THEN LET(n, n + 1)',
                                    "LET(t, CLEAN_TEXT(CELL($ws, r, 2)))",
                                    "IF IS_ERROR(CELL($ws, r, 2)) THEN LET(n, n + 1)")))

    def test_same_and_is_error_are_conditions(self):
        self.assertProblem(spec("LET(n, SAME(t, \"a\") + 1)"), "can't apply +")
        self.assertClean(spec('LET(flag, SAME(t, "a") AND NOT IS_ERROR(t))'))


class Interpreter(unittest.TestCase):
    def test_the_example_passes_its_tests(self):
        results = run_tests(TIDY)
        self.assertEqual(len(results), 2)
        self.assertTrue(all(r.passed for r in results), results)

    def test_mod_by_zero_is_a_logged_runtime_error(self):
        text = "\n".join(HEAD + ["OP       LET(n, MOD(5, n))", "OP       RETURN(n)", "ERR      standard",
                                 "FIXTURE  t_01.xlsx sheet=S cols=A:Status rows=2 data=A:text(3)",
                                 "TEST     fixture=t_01.xlsx call=P(S) assert=RESULT = -2 AND ERROR_COUNT = 1",
                                 "TEST     fixture=t_01.xlsx call=P(S) assert=COUNTA(S!A:A) = PRE_COUNTA(S!A:A)"])
        self.assertTrue(all(r.passed for r in run_tests(text)))

    def test_iif_evaluates_both_values(self):
        # The FALSE branch divides by zero, so the call fails even though the condition is TRUE.
        text = "\n".join(HEAD + ["OP       LET(n, IIF(TRUE, 1, 1 / n))", "OP       RETURN(n)", "ERR      standard",
                                 "FIXTURE  t_01.xlsx sheet=S cols=A:Status rows=2 data=A:text(3)",
                                 "TEST     fixture=t_01.xlsx call=P(S) assert=RESULT = -2 AND ERROR_COUNT = 1",
                                 "TEST     fixture=t_01.xlsx call=P(S) assert=COUNTA(S!A:A) = PRE_COUNTA(S!A:A)"])
        self.assertTrue(all(r.passed for r in run_tests(text)))


class Compiler(unittest.TestCase):
    def test_every_meaning_goes_through_forgelib(self):
        vba = compile_text(TIDY, "tidy.vcl")
        for helper in ('ForgeSame(ws.Cells(r, cStatus).Value, "Ready")', "ForgeCleanText(",
                       'ForgeParseDate(ws.Cells(r, cDue).Value, "dd/mm/yyyy")', "ForgeIif(",
                       "ForgeMod(ForgeNum(ws.Cells(r, cQty).Value), 2)", "ForgeIsError("):
            self.assertIn(helper, vba)
        self.assertNotIn(" Mod ", vba)
        self.assertNotIn("CDate(ForgeParseDate", vba)

    def test_module_lints_clean(self):
        self.assertEqual(lint(assemble([compile_text(TIDY, "tidy.vcl")])), [])

    def test_forgelib_defines_each_helper(self):
        lib = (ROOT / "vba" / "ForgeLib.bas").read_text(encoding="utf-8")
        for helper in ("ForgeSame", "ForgeCleanText", "ForgeIsError", "ForgeParseDate", "ForgeIif", "ForgeMod"):
            self.assertIn(f"Public Function {helper}(", lib)
        for banned in ("CDate(", "DateValue("):
            body = lib[lib.index("Public Function ForgeParseDate("):lib.index("End Function", lib.index("Public Function ForgeParseDate("))]
            self.assertNotIn(banned, body)


if __name__ == "__main__":
    unittest.main()
