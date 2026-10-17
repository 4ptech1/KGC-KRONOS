"""VCL 0.2 (docs/vep-0.2-core.md) through every stage: the version line, the validator, the
type checker, the interpreter, the compiler and ForgeLib.
PR A: the value functions SAME, CLEAN_TEXT, IS_ERROR, PARSE_DATE, IIF and MOD.
PR B: ForgeSetValue (S17), SKIP_IF, EXIT_LOOP_IF and WRITE_ROW."""
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
from vcl_mutate import mutate_report  # noqa: E402
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

# Rows 2 to 6: Ready (no Qty), Ready 3, Closed 9, Ready 4, End. The loop skips row 2 (no
# number) and row 4 (not Ready), writes rows 3 and 5, and leaves at row 6, so the Ready rows
# below End are never written.
LISTER = """VCL 0.2
TASK
# ASSUME: Data headers (row 1) include "Status" and "Qty"; a Status of "End" ends the list
PROC     ListReady
KIND     Function
MODULE   modList
ARGS     ws As Worksheet, out As Worksheet
RET      Long
VAR      cStatus As Long
VAR      cQty As Long
VAR      outRow As Long
PRE      COL($ws, "Status") > 0; COL($ws, "Qty") > 0
OP       LET(cStatus, COL($ws, "Status"))
OP       LET(cQty, COL($ws, "Qty"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Row", "Status", "Qty")
OP       LET(outRow, 1)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cStatus) var=r
OP         EXIT_LOOP_IF SAME(CELL($ws, r, cStatus), "End")
OP         SKIP_IF NOT SAME(CELL($ws, r, cStatus), "Ready")
OP         SKIP_IF NOT IS_NUMBER(CELL($ws, r, cQty))
OP         LET(outRow, outRow + 1)
OP         WRITE_ROW($out, outRow, r, CELL($ws, r, cStatus), CELL($ws, r, cQty) * 2)
OP       RETURN(outRow - 1)
ERR      standard
FIXTURE  list_01.xlsx sheet=Data cols=A:Status,B:Qty rows=8 data=A:values("Ready"x5,"Closed"x2,"End"x1)+boundary("Ready","Ready","Closed","Ready","End");B:int(1,9)+boundary("",3,9,4)
FIXTURE  list_01.xlsx sheet=Out cols=A:Row rows=0 stale=2
TEST     fixture=list_01.xlsx call=ListReady(Data, Out) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND COUNTA(Out!A:A) = 3 AND SUM(Out![Qty]) = 14 AND VALUE(Out!A3) = 5 AND COUNTIF(Out![Status], "Ready") = 2 AND FILL_COUNT(Out!A1:C9, #FF0000) = 0
TEST     fixture=list_01.xlsx call=ListReady(Data, Out) assert=COUNTA(Data![Status]) = PRE_COUNTA(Data![Status]) AND SUM(Data![Qty]) = PRE_SUM(Data![Qty])
TEST     fixture=list_01.xlsx call=ListReady(Out, Out) assert=RESULT = -1 AND ERROR_COUNT = 1
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



class LoopOps(unittest.TestCase):
    """PR B: SKIP_IF, EXIT_LOOP_IF and WRITE_ROW."""

    def test_the_example_is_valid_and_passes(self):
        self.assertEqual([(p.line, p.message) for p in validator.validate_text(LISTER)], [])
        results = run_tests(LISTER)
        self.assertEqual(len(results), 3)
        self.assertTrue(all(r.passed for r in results), [(r.line, r.failed) for r in results])

    def test_they_need_the_0_2_version_line(self):
        found = errors(LISTER.replace("VCL 0.2", "VCL 0.1", 1))
        for name in ("SKIP_IF", "EXIT_LOOP_IF", "WRITE_ROW"):
            self.assertIn(f"{name} is VCL 0.2: start the spec with 'VCL 0.2' to use it", found)

    def test_they_only_live_inside_a_loop(self):
        for op in ("SKIP_IF", "EXIT_LOOP_IF"):
            found = errors(LISTER.replace("OP       LET(outRow, 1)", f"OP       {op} outRow = 1"))
            self.assertIn(f"{op} is only allowed inside LOOP_ROWS, LOOP_COLS or FOR_EACH", found)
        self.assertIn("SKIP_IF can't have nested steps",
                      errors(LISTER.replace("OP         LET(outRow, outRow + 1)", "OP           LET(outRow, outRow + 1)")))
        self.assertIn("IF ... THEN needs exactly one op after THEN, like THEN LET(x, 1)",
                      errors(LISTER.replace("OP         SKIP_IF NOT IS_NUMBER", "OP         IF TRUE THEN SKIP_IF NOT IS_NUMBER")))

    def test_write_row_arguments(self):
        self.assertIn("WRITE_ROW takes 3 or more arguments before its options, got 2",
                      errors(LISTER.replace('WRITE_ROW($out, 1, "Row", "Status", "Qty")', "WRITE_ROW($out, 1)")))
        self.assertIn("WRITE_ROW has no option to=",
                      errors(LISTER.replace('"Status", "Qty")', '"Status", "Qty", to=2)')))
        self.assertEqual(errors(LISTER.replace('"Status", "Qty")', '"Status", "Qty", col=1)')), [])

    def test_write_row_headers_count_for_test_ranges(self):
        # Out![Qty] is only known because WRITE_ROW writes it in row 1.
        found = errors(LISTER.replace('"Row", "Status", "Qty")', '"Row", "Status", "Amount")'))
        self.assertTrue(any("Qty" in m for m in found), found)

    def test_skip_if_guards_the_rest_of_the_pass_only(self):
        self.assertEqual(check_text(LISTER), [])
        self.assertTrue(any("arithmetic" in m for _, m in check_text(
            LISTER.replace("OP         SKIP_IF NOT IS_NUMBER(CELL($ws, r, cQty))\n", ""))))
        moved = spec(*loop("IF flag", "  SKIP_IF NOT IS_NUMBER(CELL($ws, r, 2))", "LET(n, CELL($ws, r, 2) + 1)"))
        self.assertTrue(any("arithmetic" in m for _, m in check_text(moved)))
        inside = spec(*loop("IF flag", "  SKIP_IF NOT IS_NUMBER(CELL($ws, r, 2))", "  LET(n, CELL($ws, r, 2) + 1)"))
        self.assertEqual(check_text(inside), [])

    def test_write_row_types(self):
        self.assertTrue(any("arithmetic" in m for _, m in check_text(
            spec(*loop("WRITE_ROW($ws, r, CELL($ws, r, 2) * 2)")))))
        self.assertTrue(any("expected Sheet" in m for _, m in check_text(spec("WRITE_ROW(n, 1, 2)"))))
        self.assertTrue(any("expected a value" in m for _, m in check_text(spec("WRITE_ROW($ws, 1, $ws)"))))
        self.assertEqual(check_text(spec('WRITE_ROW($ws, 1, "a", n, TODAY, col=n + 1)')), [])

    def test_skip_if_and_exit_loop_if_differ(self):
        swapped = LISTER.replace("EXIT_LOOP_IF SAME", "SKIP_IF SAME")
        self.assertFalse(all(r.passed for r in run_tests(swapped)))

    def test_mutants_of_the_new_lines_are_caught(self):
        killed, total, survivors = mutate_report(LISTER, snapshot=True)
        self.assertGreater(total, 20)
        self.assertEqual([s for s in survivors if s[0] in (19, 20, 21, 23)], [])

    def test_compiles_to_goto_next_label_and_exit_for(self):
        vba = compile_text(LISTER, "lister.vcl")
        self.assertIn("Then Exit For", vba)
        self.assertIn("Then GoTo Next_18", vba)   # the LOOP_ROWS line
        lines = vba.splitlines()
        label = lines.index("Next_18:")
        self.assertTrue(lines[label + 1].strip().endswith("Next r"))
        self.assertIn('ForgeWriteRow out, 1, 1, Array("Row", "Status", "Qty")', vba)
        self.assertEqual(lint(assemble([vba])), [])

    def test_each_loop_gets_its_own_label(self):
        text = LISTER.replace("OP       RETURN(outRow - 1)",
                              "OP       FOR_EACH var=sh in=SHEETS(THIS_BOOK)\n"
                              "OP         SKIP_IF SAME(outRow, 0)\n"
                              "OP         LET(outRow, outRow + 0)\n"
                              "OP       RETURN(outRow - 1)")
        self.assertEqual(errors(text), [])
        vba = compile_text(text)
        self.assertIn("Next_18:", vba.splitlines())
        self.assertIn("Next_24:", vba.splitlines())
        self.assertEqual(lint(assemble([vba])), [])

    def test_no_label_without_skip_if(self):
        vba = compile_text(LISTER.replace("SKIP_IF NOT SAME", "EXIT_LOOP_IF NOT SAME")
                                 .replace("SKIP_IF NOT IS_NUMBER", "EXIT_LOOP_IF NOT IS_NUMBER"))
        self.assertNotIn("Next_18", vba)


class TextStaysText(unittest.TestCase):
    """PR B: S17, SET_VALUE, WRITE_ROW and COPY_ROW store text as exactly that text."""

    def test_writes_go_through_forge_set_value(self):
        vba = compile_text(TIDY)
        self.assertIn("ForgeSetValue ws.Cells(r, 4), ForgeCleanText(", vba)
        self.assertNotIn(".Value = ", vba.split("Public Function TidyRows")[1])

    def test_forgelib_writes_every_cell_through_the_guard(self):
        lib = (ROOT / "vba" / "ForgeLib.bas").read_text(encoding="utf-8")
        for proc in ("ForgeWriteRow", "ForgeCopyRow"):
            body = lib[lib.index(f"Public Sub {proc}("):lib.index("End Sub", lib.index(f"Public Sub {proc}("))]
            self.assertIn("ForgeSetValue ", body)
            self.assertNotIn(".Value =", body)

    def test_the_fixture_builder_uses_the_same_rule(self):
        import vcl_fixtures
        self.assertIs(vcl_fixtures.needs_text_guard, semantics.needs_text_guard)
        self.assertEqual(vcl_fixtures.vba_literal("#N/A"), '"\'#N/A"')
        self.assertEqual(vcl_fixtures.vba_literal("Short"), '"Short"')


if __name__ == "__main__":
    unittest.main()
