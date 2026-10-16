"""Unit tests for compiler/validator.py. Run: python -m unittest discover tests

Findings 1 to 7 are the starter-checker bugs from the Phase 0 review (PR #1).
"""
import glob
import os
import random
import re
import subprocess
import sys
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "compiler"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from test_parser import MINIMAL_BUILD  # noqa: E402
from validator import validate_text  # noqa: E402

MINIMAL = """VCL 0.1
TASK
# ASSUME: Data headers (row 1) include "ID" and "Status"
PROC     CountOpenRows
KIND     Function
MODULE   modDemo
ARGS     ws As Worksheet
RET      Long
VAR      cStatus As Long
VAR      found As Long
PRE      COL($ws, "Status") > 0
OP       LET(cStatus, COL($ws, "Status"))
OP       LET(found, 0)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cStatus) var=r
OP         IF UPPER(TRIM(CELL($ws, r, cStatus))) = "OPEN" THEN LET(found, found + 1)
OP       RETURN(found)
ERR      standard
FIXTURE  demo_01.xlsx sheet=Data cols=A:ID,B:Status rows=10 data=A:seq(1,1);B:values("Open"x6,"Closed"x4)
TEST     fixture=demo_01.xlsx call=CountOpenRows(Data) runs=2 assert=RESULT = PRE_COUNTIF(Data![Status], "Open") AND ERROR_COUNT = 0
TEST     fixture=demo_01.xlsx call=CountOpenRows(Data) assert=COUNTA(Data![Status]) = PRE_COUNTA(Data![Status])
"""

# Two sheets: Report gets Data's header row (COPY_ROW) plus a written "Checked" header.
WITH_OUTPUT = """VCL 0.1
TASK
PROC     ListOpenRows
KIND     Function
MODULE   modDemo
ARGS     ws As Worksheet, out As Worksheet
RET      Long
VAR      cStatus As Long
VAR      outRow As Long
PRE      COL($ws, "Status") > 0
OP       LET(cStatus, COL($ws, "Status"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       COPY_ROW($ws, 1, $out, 1, cols=2)
OP       SET_VALUE(CELL($out, 1, 3), "Checked")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cStatus) var=r
OP         IF UPPER(TRIM(CELL($ws, r, cStatus))) = "OPEN"
OP           COPY_ROW($ws, r, $out, outRow, cols=2)
OP           SET_VALUE(CELL($out, outRow, 3), "Yes")
OP           LET(outRow, outRow + 1)
OP       RETURN(outRow - 2)
ERR      standard
FIXTURE  demo_02.xlsx sheet=Data cols=A:ID,B:Status rows=10 data=A:seq(1,1);B:values("Open"x6,"Closed"x4)
FIXTURE  demo_02.xlsx sheet=Report cols=A:Note rows=0
TEST     fixture=demo_02.xlsx call=ListOpenRows(Data, Report) runs=2 assert=RESULT = PRE_COUNTIF(Data![Status], "Open") AND COUNTIF(Report![Status], "Open") = RESULT AND COUNTA(Report![Checked]) = RESULT + 1 AND ERROR_COUNT = 0
TEST     fixture=demo_02.xlsx call=ListOpenRows(Data, Report) assert=COUNTA(Data![Status]) = PRE_COUNTA(Data![Status])
TEST     fixture=demo_02.xlsx call=ListOpenRows(Report, Report) assert=RESULT = -1 AND ERROR_COUNT = 1
"""

RETURN_LINE = "OP       RETURN(found)"


def read(path):
    with open(os.path.join(ROOT, path), encoding="utf-8") as f:
        return f.read()


def edit(text, old, new):
    assert old in text, old
    return text.replace(old, new)


def with_op(step, text=MINIMAL):
    """MINIMAL with one more OP line just before RETURN."""
    return edit(text, RETURN_LINE, f"OP       {step}\n{RETURN_LINE}")


def errors(text):
    return [p.message for p in validate_text(text) if p.severity == "error"]


def warnings(text):
    return [p.message for p in validate_text(text) if p.severity == "warning"]


class Basics(unittest.TestCase):
    def test_minimal_task_is_valid(self):
        self.assertEqual(validate_text(MINIMAL), [])
        self.assertEqual(validate_text(WITH_OUTPUT), [])

    def test_question_reply_is_valid(self):
        self.assertEqual(validate_text("# QUESTION: Which folder should I clean?\n"), [])

    def test_question_reply_contains_nothing_else(self):
        self.assertIn("a QUESTION reply contains only its '# QUESTION: <question>' line",
                      errors("# QUESTION: Which folder?\nVCL 0.1\nTASK\n"))

    def test_minimal_build_is_valid(self):
        self.assertEqual(validate_text(MINIMAL_BUILD), [])


class Finding1ValidOpsAccepted(unittest.TestCase):
    """The starter checker rejected valid ops and options."""

    def test_valid_options_are_accepted(self):
        steps = [
            "CALL(modDemo.LogRun, $ws)",
            "FORMAT(CELL($ws, 1, 1), bold=yes, width=auto)",
            "FORMAT(CELL($ws, 1, 1), width=12)",
            'SET_FORMULA(CELL($ws, 1, 9), "=1+1", style=a1)',
            'SET_FORMULA(CELL($ws, 1, 9), "=R1C1", style=r1c1)',
            'SAVE_AS(THIS_BOOK, "out.xlsx", format=xlsx)',
            'SAVE_AS(THIS_BOOK, "out.xlsm", format=xlsm)',
            'SAVE_AS(THIS_BOOK, "out.csv", format=csv)',
            "FORMAT(CELL($ws, 1, 1), fill=#FFCCCC)",
            "FORMAT(CELL($ws, 1, 1), fill=#ffcccc)",
            "FORMAT(CELL($ws, 1, 1), fill=#1F4E78)",
            'FORMAT(CELL($ws, 1, 1), numfmt="0.0%", bold=no, fill=#FF0000)',
            "DEDUPE(USED_RANGE($ws), cols=1,2, header=yes)",
            "SORT(USED_RANGE($ws), key=cStatus, order=asc, header=yes)",
            'IF IS_NOTHING(FIND(USED_RANGE($ws), "x", match=whole)) THEN LET(found, 0)',
        ] + [f"FORMAT(CELL($ws, 1, 1), fill={name})"
             for name in ("none", "black", "white", "red", "green", "blue", "yellow", "magenta", "cyan")]
        for step in steps:
            with self.subTest(step=step):
                self.assertEqual(errors(with_op(step)), [])

    def test_invalid_option_values_are_errors(self):
        cases = {
            "FORMAT(CELL($ws, 1, 1), fill=pink)": "fill= must be",
            "FORMAT(CELL($ws, 1, 1), fill=#FFCC)": "fill= must be",
            "FORMAT(CELL($ws, 1, 1), width=wide)": "width= must be auto or a number",
            'SET_FORMULA(CELL($ws, 1, 9), "=1", style=b2)': "style= must be a1|r1c1",
            'SAVE_AS(THIS_BOOK, "out.pdf", format=pdf)': "format= must be xlsx|xlsm|csv",
            "FORMAT(CELL($ws, 1, 1), colour=red)": "FORMAT has no option colour=",
            "DICT_ADD(found, 1, 2)": "DICT_ADD needs mode=",
            "SORT(USED_RANGE($ws), key=1, order=up)": "order= must be asc|desc",
            "CLEAR(USED_RANGE($ws), what=everything)": "what= must be all|contents|formats",
        }
        for step, message in cases.items():
            with self.subTest(step=step):
                self.assertTrue(any(message in e for e in errors(with_op(step))), errors(with_op(step)))

    def test_call_target_must_be_module_dot_proc(self):
        self.assertTrue(any("CALL needs <module>.<Proc>" in e for e in errors(with_op("CALL(LogRun, $ws)"))))


class Finding2BadLinesReportedNotCrashing(unittest.TestCase):
    """The starter checker crashed on these lines instead of reporting them."""

    def test_lowercase_as(self):
        problems = validate_text(edit(MINIMAL, "VAR      found As Long", "VAR      found as Long"))
        self.assertIn((10, "VAR: expected '<name> As <Type>', got 'found as Long'"),
                      [(p.line, p.message) for p in problems])

    def test_op_not_starting_with_a_capital_op(self):
        problems = validate_text(edit(MINIMAL, "OP       LET(found, 0)", "OP       let(found, 0)"))
        self.assertIn((13, "unknown op let"), [(p.line, p.message) for p in problems])

    def test_fixture_without_sheet(self):
        problems = validate_text(edit(MINIMAL, "demo_01.xlsx sheet=Data cols=", "demo_01.xlsx cols="))
        self.assertTrue(any(p.line == 18 and "missing sheet=" in p.message for p in problems))

    def test_damaged_specs_never_crash(self):
        rng = random.Random(11)
        texts = [read(path) for path in sorted(glob.glob(os.path.join(ROOT, "specs", "*.vcl")))]
        texts += [MINIMAL, WITH_OUTPUT, MINIMAL_BUILD]
        pool = ' ()[],;=<>"$#!:+-*/&.\nabcXYZ019_'
        for _ in range(400):
            chars = list(rng.choice(texts))
            for _ in range(rng.randint(1, 6)):
                at = rng.randrange(len(chars))
                if rng.random() < 0.5:
                    del chars[at]
                else:
                    chars.insert(at, rng.choice(pool))
            for p in validate_text("".join(chars)):
                self.assertIn(p.severity, ("error", "warning"))
                self.assertIsInstance(p.line, int)


class Finding3HeadersPerSheet(unittest.TestCase):
    """A TEST [Header] must exist on that exact sheet, not on any sheet."""

    def test_header_from_another_sheet_is_an_error(self):
        text = edit(read("specs/task_01.vcl"), 'COUNTIF(Shortages![Remark], "Required qty missing") = 3',
                    'COUNTIF(Shortages![Required Qty], ">0") = 3')
        self.assertTrue(any("sheet Shortages has no [Required Qty] header" in e for e in errors(text)))

    def test_written_and_copied_headers_count_only_on_their_sheet(self):
        self.assertEqual(errors(WITH_OUTPUT), [])
        wrong = edit(WITH_OUTPUT, "COUNTA(Data![Status]) = PRE_COUNTA", "COUNTA(Data![Checked]) = PRE_COUNTA")
        self.assertTrue(any("sheet Data has no [Checked] header" in e for e in errors(wrong)))

    def test_copied_header_row_is_needed(self):
        text = edit(read("specs/task_07.vcl"), "OP       COPY_ROW($itp, 1, $out, 1, cols=lastCol)\n", "")
        self.assertTrue(any("sheet CallNotice has no [Inspection Status] header" in e for e in errors(text)))

    def test_written_header_is_needed(self):
        text = edit(read("specs/task_10.vcl"), 'OP       SET_VALUE(CELL($out, 1, lastCol + 1), "Status")\n', "")
        self.assertTrue(any("sheet Overdue has no [Status] header" in e for e in errors(text)))

    def test_unknown_sheet_is_an_error(self):
        text = edit(MINIMAL, "COUNTA(Data![Status]) =", "COUNTA(Orders![Status]) =")
        self.assertIn("assertion: sheet Orders is not in demo_01.xlsx", errors(text))


class Finding4PreIsChecked(unittest.TestCase):
    def test_unknown_function_and_names_in_pre(self):
        text = edit(MINIMAL, 'PRE      COL($ws, "Status") > 0',
                    'PRE      COL($ws, "Status") > 0; XLOOKUP($nope, bogus) > 0')
        found = errors(text)
        for message in ("unknown function XLOOKUP", "undeclared argument $nope", "undeclared name bogus"):
            self.assertIn(message, found)

    def test_vba_in_pre(self):
        text = edit(MINIMAL, 'PRE      COL($ws, "Status") > 0', 'PRE      Cells(1, 1).Value <> ""')
        self.assertTrue(any(e.startswith("VBA syntax in spec") for e in errors(text)))


class Finding5ArgumentsNeedDollar(unittest.TestCase):
    def test_argument_without_dollar(self):
        text = edit(MINIMAL, "CELL($ws, r, cStatus)", "CELL(ws, r, cStatus)")
        self.assertIn("argument ws must be written $ws", errors(text))

    def test_variable_with_dollar(self):
        text = edit(MINIMAL, "LET(found, found + 1)", "LET(found, $found + 1)")
        self.assertIn("found is a variable, not an argument: write it without $", errors(text))

    def test_undeclared_names(self):
        self.assertIn("undeclared argument $sheet", errors(with_op("LET(found, LASTROW($sheet, 1))")))
        self.assertIn("undeclared name total", errors(with_op("LET(found, total)")))


class Finding6NoFilesFails(unittest.TestCase):
    def run_validator(self, *args):
        return subprocess.run([sys.executable, os.path.join("compiler", "validator.py"), *args],
                              cwd=ROOT, capture_output=True, text=True)

    def test_no_files_exits_1(self):
        result = self.run_validator()
        self.assertEqual(result.returncode, 1)
        self.assertIn("no spec files given", result.stderr)

    def test_missing_file_exits_1(self):
        result = self.run_validator("specs/no_such_spec.vcl")
        self.assertEqual(result.returncode, 1)
        self.assertIn("specs/no_such_spec.vcl:1: error: cannot read file", result.stdout)

    def test_wildcards_work_without_a_shell(self):
        result = self.run_validator("specs/*.vcl")
        self.assertEqual(result.returncode, 0, result.stdout)
        self.assertIn("10 files checked: 0 errors, 1 warning", result.stdout)


class Finding7KnownGaps(unittest.TestCase):
    def test_missing_required_keys(self):
        cases = {"PROC": "PROC     CountOpenRows\n", "MODULE": "MODULE   modDemo\n", "ERR": "ERR      standard\n",
                 "KIND": "KIND     Function\n"}
        for key, line in cases.items():
            with self.subTest(key=key):
                self.assertIn(f"missing {key}", errors(edit(MINIMAL, line, "")))
        no_ops = "\n".join(line for line in MINIMAL.splitlines() if not line.startswith("OP"))
        self.assertIn("missing OP", errors(no_ops))

    def test_duplicate_keys(self):
        text = edit(MINIMAL, "ERR      standard", "ERR      standard\nERR      standard")
        self.assertIn("ERR appears more than once", errors(text))

    def test_perf_allow_and_err_values(self):
        perf = edit(MINIMAL, "ERR      standard", "PERF     screen_off turbo\nERR      standard")
        self.assertIn("unknown PERF value turbo (use screen_off, calc_manual, events_off)", errors(perf))
        allow = edit(MINIMAL, "ERR      standard", "ERR      standard\nALLOW    reboot")
        self.assertIn("unknown ALLOW value reboot (use file_delete, file_overwrite, shell, http)", errors(allow))
        err = edit(MINIMAL, "ERR      standard", "ERR      custom")
        self.assertIn("ERR must be 'standard', got 'custom'", errors(err))

    def test_nesting_needs_a_block_op(self):
        text = edit(MINIMAL, "OP       LET(found, 0)", "OP       LET(found, 0)\nOP         LET(found, 1)")
        self.assertIn("LET can't have nested steps", errors(text))
        empty_if = with_op("IF found > 0")
        self.assertIn("IF has no nested steps under it", errors(empty_if))

    def test_else_must_follow_a_block_if(self):
        self.assertIn("ELSE must come right after a block IF at the same level",
                      errors(edit(MINIMAL, RETURN_LINE, "OP       ELSE\nOP         LET(found, 0)\n" + RETURN_LINE)))

    def test_build_structure(self):
        no_task = MINIMAL_BUILD[:MINIMAL_BUILD.rindex("---")]
        self.assertIn("SIG modFormat.BoldHeader has no TASK section after ---", errors(no_task))
        self.assertIn("ENTRY must be modMain.Run, got 'modMain.Start'",
                      errors(edit(MINIMAL_BUILD, "ENTRY    modMain.Run", "ENTRY    modMain.Start")))
        self.assertIn("CALL modFormat.Missing has no SIG in this BUILD",
                      errors(edit(MINIMAL_BUILD, "CALL(modFormat.BoldHeader, $ws)", "CALL(modFormat.Missing, $ws)")))
        self.assertIn("TASK modFormat.BoldHeader ARGS don't match its SIG",
                      errors(edit(MINIMAL_BUILD, "ARGS     ws As Worksheet\nOP       FORMAT",
                                  "ARGS     ws As Range\nOP       FORMAT")))


class Revision3(unittest.TestCase):
    def test_fixture_target_is_an_error(self):
        text = edit(MINIMAL, "fixture=demo_01.xlsx call=CountOpenRows(Data) assert",
                    "fixture=target call=CountOpenRows(Data) assert")
        self.assertIn("fixture=target is not allowed: use a named synthetic FIXTURE (rule 18)", errors(text))

    def test_countifs_and_sumifs_are_accepted(self):
        text = edit(MINIMAL, 'RESULT = PRE_COUNTIF(Data![Status], "Open")',
                    'RESULT = PRE_COUNTIFS(Data![Status], "Open", Data![ID], ">0") '
                    'AND SUMIFS(Data![ID], Data![Status], "<>Closed", Data![ID], "<>") >= 0')
        self.assertEqual(errors(text), [])

    def test_countifs_and_sumifs_argument_counts(self):
        bad_countifs = edit(MINIMAL, 'PRE_COUNTIF(Data![Status], "Open")',
                            'PRE_COUNTIFS(Data![Status], "Open", Data![ID])')
        self.assertIn("assertion: PRE_COUNTIFS takes range, criteria pairs, got 3", errors(bad_countifs))
        bad_sumifs = edit(MINIMAL, 'PRE_COUNTIF(Data![Status], "Open")', 'SUMIFS(Data![ID], Data![Status])')
        self.assertIn("assertion: SUMIFS takes a sum range, then range, criteria pairs, got 2", errors(bad_sumifs))

    def test_unknown_assertion_function(self):
        text = edit(MINIMAL, 'PRE_COUNTIF(Data![Status], "Open")', 'XLOOKUP(Data![Status], "Open")')
        self.assertIn("assertion: unknown function XLOOKUP", errors(text))


class GrammarAndTypeStages(unittest.TestCase):
    """The validator also runs vcl/grammar.lark and the type checker (revision 4)."""

    def test_type_checker_catches_unguarded_values(self):
        self.assertIn("type: raw cell value CELL($ws,2,cStatus) used in arithmetic: guard it with IS_NUMBER "
                      "or IS_DATE first", errors(with_op("LET(found, CELL($ws, 2, cStatus) + 1)")))
        flag = edit(MINIMAL, "VAR      found As Long", "VAR      found As Long\nVAR      ok As Boolean")
        flag = edit(flag, RETURN_LINE, "OP       LET(ok, IS_NUMBER(CELL($ws, 2, 1)))\nOP       IF ok\n"
                                       "OP         LET(found, CELL($ws, 2, 1))\n" + RETURN_LINE)
        self.assertTrue(any(e.startswith("type: assigning raw cell value") for e in errors(flag)), errors(flag))

    def test_grammar_catches_what_rule_checks_miss(self):
        self.assertIn("grammar: syntax error at column 14 near '<'",
                      errors(with_op("IF 1 < found < 3 THEN LET(found, 1)")))           # G7
        star = edit(MINIMAL, "assert=RESULT = PRE_COUNTIF", "assert=RESULT * 1 = PRE_COUNTIF")
        self.assertTrue(any(e.startswith("grammar: syntax error") for e in errors(star)), errors(star))

    def test_one_message_per_broken_line(self):
        # The type checker would add "argument ws must be written $ws" and "raw cell value ... used in
        # arithmetic" for this line too; stage 1 already reports it, so they are left out.
        self.assertEqual(errors(with_op("LET(found, CELL(ws, 2, 1) + 1)")), ["argument ws must be written $ws"])

    def test_grammar_line_numbers_inside_a_build_are_the_files(self):
        bad = edit(MINIMAL_BUILD, "OP       LET(found, 0)", "OP       LET(found, 1 < 2 < 3)")
        line = bad.splitlines().index("OP       LET(found, 1 < 2 < 3)") + 1
        self.assertIn(line, [p.line for p in validate_text(bad) if p.message.startswith("grammar:")])

    def test_byte_order_mark_is_fine(self):
        self.assertEqual(validate_text("\ufeff" + MINIMAL), [])

    def test_doubled_quotes_escape_a_quote(self):  # G6
        self.assertEqual(errors(with_op('LOG("Say ""hi""")')), [])
        self.assertEqual(errors(edit(MINIMAL, '"Closed"x4', '"Clo""sed"x4')), [])
        self.assertEqual(errors(edit(MINIMAL, 'PRE_COUNTIF(Data![Status], "Open")',
                                     'PRE_COUNTIF(Data![Status], "Open") + COUNTIF(Data![Status], "a""b")')), [])

    def test_note_as_a_one_line_if_action(self):
        self.assertEqual(errors(with_op('IF found > 0 THEN NOTE "flag the row for review"')), [])

    def test_reserved_option_words(self):  # G2
        self.assertIn("key is a reserved option word: choose another name (G2)",
                      errors(edit(MINIMAL, "VAR      found As Long", "VAR      found As Long\nVAR      key As Long")))
        loop = edit(MINIMAL, "var=r", "var=name").replace("CELL($ws, r, cStatus)", "CELL($ws, name, cStatus)")
        self.assertIn("name is a reserved option word: choose another loop variable name (G2)", errors(loop))
        # only reserved words make options, so a declared name before "=" is a comparison
        self.assertEqual(errors(with_op("SET_VALUE(CELL($ws, 1, 3), found=0)")), [])
        self.assertIn("FORMAT has no option colour=", errors(with_op("FORMAT(CELL($ws, 1, 1), colour=red)")))


class Warnings(unittest.TestCase):
    def test_more_than_40_op_lines_warns(self):
        extra = "OP       LET(found, found + 1)\n"
        forty = edit(MINIMAL, RETURN_LINE, extra * 35 + RETURN_LINE)  # 5 + 35 = 40 OP lines
        self.assertEqual(validate_text(forty), [])
        forty_one = edit(MINIMAL, RETURN_LINE, extra * 36 + RETURN_LINE)
        self.assertEqual(errors(forty_one), [])
        self.assertEqual(warnings(forty_one),
                         ["TASK has 41 OP lines (more than 40): consider splitting it into a BUILD (rule 9)"])

    def test_allow_warns(self):
        text = edit(MINIMAL, "ERR      standard", "ERR      standard\nALLOW    file_delete")
        self.assertEqual(errors(text), [])
        self.assertEqual(warnings(text),
                         ["ALLOW file_delete: risky actions can't be tested on the work PC before March 2027"])


class OtherRules(unittest.TestCase):
    def test_delete_row_only_inside_loop_rows_dir_up(self):
        self.assertIn("DELETE_ROW is only allowed inside LOOP_ROWS dir=up (rule 10)",
                      errors(with_op("DELETE_ROW($ws, 2)")))
        up = edit(MINIMAL, "LOOP_ROWS dir=down from=2 to=LASTROW($ws, cStatus)",
                  "LOOP_ROWS dir=up from=LASTROW($ws, cStatus) to=2")
        up = edit(up, "THEN LET(found, found + 1)", "THEN DELETE_ROW($ws, r)")
        self.assertEqual(errors(up), [])

    def test_wrong_sheet_test_is_required(self):
        text = edit(WITH_OUTPUT, "TEST     fixture=demo_02.xlsx call=ListOpenRows(Report, Report) "
                                 "assert=RESULT = -1 AND ERROR_COUNT = 1\n", "")
        self.assertTrue(any("no wrong-sheet TEST expects RESULT = -1 AND ERROR_COUNT = 1" in e for e in errors(text)))

    def test_test_that_fails_pre_must_expect_minus_1(self):
        text = edit(WITH_OUTPUT, "call=ListOpenRows(Report, Report) assert=RESULT = -1 AND ERROR_COUNT = 1",
                    "call=ListOpenRows(Report, Report) assert=ERROR_COUNT = 0")
        self.assertIn('TEST fails PRE (sheet Report has no "Status" header) but doesn\'t expect RESULT = -1',
                      errors(text))

    def test_header_used_in_op_is_checked_in_pre(self):
        text = with_op('LET(found, COL($ws, "ID"))')
        self.assertIn('header "ID" on $ws is used but not checked in PRE (rule 14)', errors(text))

    def test_fixture_rules(self):
        self.assertIn('B: values(...) counts add up to 9, but rows=10',
                      errors(edit(MINIMAL, '"Closed"x4', '"Closed"x3')))
        self.assertIn("FIXTURE cols must run A, B in order, got A, C",
                      errors(edit(MINIMAL, "cols=A:ID,B:Status", "cols=A:ID,C:Status")))
        self.assertIn("fixture file Demo.xlsx must be named like snake_case_01.xlsx (rule 12)",
                      errors(MINIMAL.replace("demo_01.xlsx", "Demo.xlsx")))

    def test_names(self):
        self.assertIn("name Found must be camelCase (rule 12)", errors(MINIMAL.replace("found", "Found")))
        self.assertIn("loop variable r is declared automatically: don't declare it in ARGS or VAR",
                      errors(edit(MINIMAL, "VAR      found As Long", "VAR      found As Long\nVAR      r As Long")))


class RealSpecs(unittest.TestCase):
    def test_all_ten_specs(self):
        paths = sorted(glob.glob(os.path.join(ROOT, "specs", "task_*.vcl")))
        self.assertEqual(len(paths), 10)
        for path in paths:
            name = os.path.basename(path)
            with self.subTest(spec=name):
                problems = validate_text(read(path))
                self.assertEqual([p.message for p in problems if p.severity == "error"], [])
                found = [(p.line, p.message) for p in problems if p.severity == "warning"]
                if name == "task_04.vcl":
                    self.assertEqual(found, [(28, "TASK has 52 OP lines (more than 40): consider splitting it "
                                                  "into a BUILD (rule 9)")])
                else:
                    self.assertEqual(found, [])

    def test_orchestrator_example_1(self):
        prompt = read("vcl/orchestrator_prompt.txt")
        example = prompt[prompt.index("EXAMPLE 1"):prompt.index("EXAMPLE 2")]
        spec = example[example.index("VCL 0.1"):]
        self.assertEqual(validate_text(spec), [])

    def test_example_1_file_is_the_prompts_example_1(self):
        prompt = read("vcl/orchestrator_prompt.txt")
        example = prompt[prompt.index("EXAMPLE 1"):prompt.index("EXAMPLE 2")]
        self.assertEqual(read("vcl/examples/example1_rev3.vcl").strip(), example[example.index("VCL 0.1"):].strip())

    def test_spec_md_and_prompt_share_the_reference_text(self):
        spec, prompt = read("vcl/spec.md"), read("vcl/orchestrator_prompt.txt")

        def between(text, start, end):
            match = re.search(re.escape(start) + r".*?" + re.escape(end), text, re.S)
            self.assertIsNotNone(match, start)
            return match.group(0)

        for start, end in [("1.  Reply with exactly", "real workbooks."),
                           ("TASK FORMAT", "the same spec always builds the same file.")]:
            self.assertEqual(between(spec, start, end), between(prompt, start, end))


if __name__ == "__main__":
    unittest.main()
