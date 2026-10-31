"""VEP 0.5 (docs/vep-0.5-build.md): CALL and BUILD specs. Phase 1: the VCL 0.5 line, CALL into=,
and the validator and type-checker rules that let a BUILD's TESTs prove the whole tool (D49, D50).
Phase 2: the interpreter runs BUILD TESTs (D52). Phase 3: one module per BUILD, and mutation (D53)."""
import contextlib
import datetime
import io
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
import validator  # noqa: E402
from vcl_samples import old_example  # noqa: E402
from vcl_interp import VclError, run_tests  # noqa: E402
from vba_lint import lint  # noqa: E402
from vcl_codegen import CompileError, compile_build  # noqa: E402
from vcl_mutate import mutate_report, untested_lines  # noqa: E402
import vclc  # noqa: E402

# Example 4 (D58) is a small tool: RunHeadless marks old open NCRs Overdue (modSteps.MarkOverdue)
# and rebuilds a summary of the statuses (modSteps.WriteSummary); modMain.Run starts it on the
# real sheets. The prompt's copy of it is kept equal to this file by tests/test_validator.py.
NCR_REVIEW = old_example("example4_build.vcl")   # as the VCL 0.5 BUILD it was before its GOAL lines (D63)


# Phase 4a (D55): ROW_OF finds the row of a FIND result; DEDUPE into= counts the rows it removed.
STOCK = r"""VCL 0.5
TASK
# ASSUME: Stock headers (row 1) include "Part No" and "Status"
# ASSUME: a repeated Part No is a duplicate row; the first one stays
PROC     CheckPart
KIND     Function
MODULE   modStock
ARGS     ws As Worksheet, part As String
RET      Long
VAR      cPart As Long
VAR      cStatus As Long
VAR      hit As Range
VAR      removed As Long
PRE      COL($ws, "Part No") > 0; COL($ws, "Status") > 0
MODIFIES $ws
OP       LET(cPart, COL($ws, "Part No"))
OP       LET(cStatus, COL($ws, "Status"))
OP       DEDUPE(USED_RANGE($ws), cols=[cPart], header=yes, into=removed)
OP       LET(hit, FIND(USED_RANGE($ws), $part))
OP       IF NOT IS_NOTHING(hit)
OP         SET_VALUE(CELL($ws, ROW_OF(hit), cStatus), "Checked")
OP       RETURN(removed)
ERR      standard
FIXTURE  stock_01.xlsx sheet=Stock cols=A:Part No,B:Status rows=6 data=A:values("P-101"x2,"P-104"x2,"P-200"x1,""x1)+boundary("P-101","P-104","P-101","P-200","P-104","");B:values("New"x6)
FIXTURE  stock_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(6)
TEST     fixture=stock_01.xlsx call=CheckPart(Stock, "P-104") runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTA(Stock![Part No]) = 4 AND VALUE(Stock![Status]3) = "Checked" AND COUNTIF(Stock![Status], "Checked") = 1
TEST     fixture=stock_01.xlsx call=CheckPart(Stock, "P-104") assert=RESULT = 2 AND COUNTA(Stock![Status]) = PRE_COUNTA(Stock![Status]) - 2 AND VALUE(Stock![Status]2) = "New"
TEST     fixture=stock_01.xlsx call=CheckPart(Stock, "P-999") assert=RESULT = 2 AND COUNTIF(Stock![Status], "Checked") = 0
TEST     fixture=stock_01.xlsx call=CheckPart(Notes, "P-104") assert=RESULT = -1 AND ERROR_COUNT = 1
"""


# Phase 4b (D56): a header the spec writes to row 1 itself counts as checked for rule 14.
CHECKED = r"""VCL 0.5
TASK
# ASSUME: Stock headers (row 1) include "Part No"; the macro adds a "Checked" column after the last header
PROC     MarkChecked
KIND     Function
MODULE   modStock
ARGS     ws As Worksheet
RET      Long
VAR      cPart As Long
VAR      cNew As Long
VAR      n As Long
PRE      COL($ws, "Part No") > 0
MODIFIES $ws
OP       LET(cPart, COL($ws, "Part No"))
OP       LET(cNew, COL($ws, "Checked"))
OP       IF cNew = 0
OP         LET(cNew, LASTCOL($ws, 1) + 1)
OP         SET_VALUE(CELL($ws, 1, cNew), "Checked")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cPart) var=r
OP         IF NOT IS_EMPTY(CELL($ws, r, cPart))
OP           SET_VALUE(CELL($ws, r, cNew), "Yes")
OP           LET(n, n + 1)
OP       RETURN(n)
ERR      standard
FIXTURE  stock_01.xlsx sheet=Stock cols=A:Part No,B:Qty rows=4 data=A:values("P-1"x1,"P-2"x1,"P-3"x1,""x1)+boundary("P-1","","P-2","P-3");B:values("4"x2,""x1,"7"x1)+boundary("","4","4","7")
FIXTURE  stock_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(6)
TEST     fixture=stock_01.xlsx call=MarkChecked(Stock) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND VALUE(Stock!C1) = "Checked" AND COUNTIF(Stock![Checked], "Yes") = 3
TEST     fixture=stock_01.xlsx call=MarkChecked(Stock) assert=RESULT = 3 AND COUNTA(Stock![Qty]) = PRE_COUNTA(Stock![Qty])
TEST     fixture=stock_01.xlsx call=MarkChecked(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
"""


# Phase 4c (D57): an assertion names a sheet whose name has spaces in quotes, as Excel does.
OPEN_NCRS = r"""VCL 0.5
TASK
# ASSUME: the count goes to a new sheet "Open NCRs" after the register; an existing one is left as it is
# ASSUME: an NCR is open when its Status is SAME as "Open"
PROC     AddOpenNcrs
KIND     Function
MODULE   modSummary
ARGS     book As Workbook, ncr As Worksheet
RET      Long
VAR      out As Worksheet
VAR      n As Long
PRE      COL($ncr, "Status") > 0
MODIFIES $book sheets
OP       IF SHEET_EXISTS($book, "Open NCRs") THEN RETURN(0)
OP       ADD_SHEET($book, name="Open NCRs", after=$ncr, into=out)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, COL($ncr, "Status")) var=r
OP         IF SAME(CELL($ncr, r, COL($ncr, "Status")), "Open") THEN LET(n, n + 1)
OP       WRITE_ROW(out, 1, "Open NCRs", n)
OP       RETURN(n)
ERR      standard
FIXTURE  ncr_01.xlsx sheet=Ncr cols=A:NCR No,B:Status rows=6 data=A:seq(1,1);B:values("Open"x2," open "x1,"Closed"x2,""x1)+boundary("Open")
FIXTURE  ncr_01.xlsx sheet=Other cols=A:Item rows=2 data=A:seq(1,1)
TEST     fixture=ncr_01.xlsx call=AddOpenNcrs(ncr_01.xlsx, Ncr) assert=RESULT = 3 AND ERROR_COUNT = 0 AND SHEET_EXISTS("Open NCRs") AND VALUE('Open NCRs'!A1) = "Open NCRs" AND VALUE('Open NCRs'!B1) = 3 AND ROWS_USED('Open NCRs') = 1 AND COUNTA(Ncr!B:B) = PRE_COUNTA(Ncr!B:B)
TEST     fixture=ncr_01.xlsx call=AddOpenNcrs(ncr_01.xlsx, Ncr) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE('Open NCRs'!B1) = 3
TEST     fixture=ncr_01.xlsx call=AddOpenNcrs(ncr_01.xlsx, Other) assert=RESULT = -1 AND ERROR_COUNT = 1
"""


def problems(text):
    return [(p.severity, p.message) for p in validator.validate_text(text)]


def errors(text):
    return [m for s, m in problems(text) if s == "error"]


def changed(old, new, text=NCR_REVIEW, count=1):
    """The sample with one change, which must be there."""
    assert text.count(old) >= count, old
    return text.replace(old, new, count)


class VersionLine(unittest.TestCase):
    def test_the_sample_is_valid(self):
        self.assertEqual(problems(NCR_REVIEW), [])

    def test_call_into_needs_0_5(self):
        older = NCR_REVIEW.replace("VCL 0.5", "VCL 0.4")
        self.assertIn("CALL into= is VCL 0.5: start the spec with 'VCL 0.5' to use it", errors(older))

    def test_every_spec_and_example_gives_the_same_results_under_0_5(self):   # E13
        texts = [(p.name, p.read_text(encoding="utf-8")) for p in sorted((ROOT / "specs").glob("*.vcl"))]
        texts += [(name, old_example(name)) for name in ("example1_rev3.vcl", "example3_files.vcl")]   # as 0.5, before GOAL (D63)
        day = datetime.date(2026, 10, 21)
        for name, text in texts:
            first, rest = text.split("\n", 1)
            with self.subTest(spec=name):
                newer = "VCL 0.5\n" + rest
                self.assertEqual(problems(text), problems(newer))
                self.assertEqual(*[[(r.line, r.passed, r.result, r.errors, r.state, r.files, r.hidden, r.texts)
                                    for r in run_tests(t, day)] for t in (text, newer)])


class Calls(unittest.TestCase):
    def test_a_task_spec_has_nothing_to_call(self):
        task = (ROOT / "specs" / "task_02.vcl").read_text(encoding="utf-8").replace("VCL 0.2", "VCL 0.5", 1)
        task = task.replace("\nERR ", "\nOP       CALL(modOther.Helper, $ws)\nERR ", 1)
        self.assertIn("CALL runs another procedure of the same BUILD: a TASK spec has none to call (D49)",
                      errors(task))

    def test_into_takes_a_function_result_of_its_type(self):
        sub = changed("SIG      modSteps.MarkOverdue(ncr As Worksheet, limit As Long) As Long",
                      "SIG      modSteps.MarkOverdue(ncr As Worksheet, limit As Long)")
        sub = changed("PROC     MarkOverdue\nKIND     Function", "PROC     MarkOverdue\nKIND     Sub", sub)
        sub = changed("RET      Long\nVAR      cStatus As Long\nVAR      cDays", "VAR      cStatus As Long\nVAR      cDays", sub)
        sub = changed("OP       RETURN(n)\nERR      standard\n---\nVCL 0.5\nTASK\nPROC     WriteSummary",
                      "ERR      standard\n---\nVCL 0.5\nTASK\nPROC     WriteSummary", sub)
        self.assertIn("CALL modSteps.MarkOverdue into=: modSteps.MarkOverdue is a Sub, so it gives no result", errors(sub))
        text = changed("VAR      marked As Long", "VAR      marked As String")
        self.assertIn("CALL modSteps.MarkOverdue gives a Long, but VAR marked is String", errors(text))

    def test_values_fit_their_parameters(self):
        found = errors(changed("CALL(modSteps.MarkOverdue, $ncr, 30, into=marked)",
                               'CALL(modSteps.MarkOverdue, $ncr, "thirty", into=marked)'))
        self.assertIn("type: passing Text to Long parameter limit", found)
        found = errors(changed("CALL(modSteps.MarkOverdue, $ncr, 30, into=marked)",
                               "CALL(modSteps.MarkOverdue, $ncr, CELL($ncr, 2, 3), into=marked)"))
        self.assertIn("type: passing raw cell value CELL($ncr,2,3) to Long parameter limit: guard it first", found)

    def test_sheets_pass_as_the_callers_own_arguments(self):
        found = errors(changed("CALL(modSteps.WriteSummary, $ncr, $out)",
                               'CALL(modSteps.WriteSummary, $ncr, SHEET(THIS_BOOK, "Summary"))'))
        self.assertIn("CALL modSteps.WriteSummary: pass $out as one of this procedure's own $arguments, written as "
                      "$name, so MODIFIES can follow it (D49)", found)


class ModifiesFollowsTheArguments(unittest.TestCase):
    def test_the_caller_declares_what_its_callees_change(self):
        found = errors(changed("MODIFIES $ncr[\"Status\"]\nMODIFIES $out\nOP       CALL", "MODIFIES $out\nOP       CALL"))
        self.assertIn('CALL modSteps.MarkOverdue may change $ncr (MODIFIES $ncr["Status"]), so this procedure '
                      'needs MODIFIES $ncr["Status"] too (D49)', found)
        found = errors(changed("MODIFIES $out\nOP       CALL", "MODIFIES $out fills\nOP       CALL"))
        self.assertIn("CALL modSteps.WriteSummary may change $out (MODIFIES $out), so this procedure needs "
                      "MODIFIES $out too (D49)", found)

    def test_every_task_of_a_build_declares_its_writes(self):
        found = errors(changed("MODIFIES $out\nOP       LET(cStatus", "OP       LET(cStatus"))
        self.assertIn("a TASK of a BUILD declares every sheet it changes with MODIFIES, so the BUILD's TESTs can "
                      "check that nothing else changed (D49)", found)


class Structure(unittest.TestCase):
    def test_no_recursion(self):
        text = changed("OP       RETURN(n)\nERR      standard\n---\nVCL 0.5\nTASK\nPROC     WriteSummary",
                       "OP       CALL(modMain.RunHeadless, $ncr, $ncr, into=n)\nOP       RETURN(n)\nERR      standard"
                       "\n---\nVCL 0.5\nTASK\nPROC     WriteSummary")
        found = errors(text)
        self.assertTrue(any(m.startswith("CALL makes a cycle (") for m in found), found)

    def test_flow_lists_exactly_the_calls(self):
        found = errors(changed("FLOW     modMain.RunHeadless -> modSteps.WriteSummary\n", ""))
        self.assertIn("modMain.RunHeadless CALLs modSteps.WriteSummary, but no FLOW line shows "
                      "modMain.RunHeadless -> modSteps.WriteSummary (D49)", found)
        found = errors(changed("FLOW     modMain.RunHeadless -> modSteps.WriteSummary",
                               "FLOW     modMain.RunHeadless -> modSteps.WriteSummary -> modSteps.MarkOverdue"))
        self.assertIn("FLOW says modSteps.WriteSummary -> modSteps.MarkOverdue, but modSteps.WriteSummary has no "
                      "CALL to modSteps.MarkOverdue (D49)", found)

    def test_procedure_names_are_unique(self):
        text = changed("SIG      modSteps.WriteSummary(", "SIG      modSteps.RunHeadless(")
        self.assertTrue(any("modMain.RunHeadless has the same name" in m for m in errors(text)))

    def test_form_and_xlam_wait(self):
        found = errors(changed("SIG      modMain.Run()", "FORM     frmMain CONTROLS btnRun:CommandButton\nSIG      modMain.Run()"))
        self.assertTrue(any(m.startswith("FORM frmMain: a UserForm can't be pasted in as text") for m in found), found)
        found = errors(changed("TARGET   xlsm", "TARGET   xlam"))
        self.assertIn("TARGET xlam waits (D49): an add-in's THIS_BOOK is the add-in, not the data; use TARGET xlsm",
                      found)


WRITE_HEADER = 'OP       WRITE_ROW($out, 1, "Status", "Count")\n'
MARK_DAYS = 'OP       LET(cDays, COL($ncr, "Days Open"))\n'


class WholeTool(unittest.TestCase):   # D67: rules a TASK spec has, held across the TASKs of a BUILD
    def test_set_formula_and_row_moves_in_two_tasks(self):   # D43
        text = changed(WRITE_HEADER, WRITE_HEADER + 'OP       SET_FORMULA(CELL($out, 1, 3), "=COUNTA(A:A)")\n')
        text = changed(MARK_DAYS, MARK_DAYS + "OP       DEDUPE(USED_RANGE($ncr), cols=[1], header=yes)\n", text)
        self.assertIn("modSteps.WriteSummary uses SET_FORMULA and modSteps.MarkOverdue uses DEDUPE: a BUILD may not "
                      "do both, because Excel rewrites formula references when rows move, which a TEST can't prove "
                      "(D43)", errors(text))

    def test_filter_and_sort_in_two_tasks(self):   # D42
        text = changed(MARK_DAYS, MARK_DAYS + 'OP       FILTER(USED_RANGE($ncr), col=cStatus, crit="Open")\n')
        text = changed(WRITE_HEADER, WRITE_HEADER + "OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)\n",
                       text)
        self.assertIn("modSteps.MarkOverdue uses FILTER and modSteps.WriteSummary uses SORT: a BUILD may not do both, "
                      "because how Excel sorts rows that FILTER hid can't be proven by a TEST (D42)", errors(text))

    def test_both_in_one_task_is_that_tasks_own_error(self):
        found = errors(changed(WRITE_HEADER, WRITE_HEADER + 'OP       SET_FORMULA(CELL($out, 1, 3), "=COUNTA(A:A)")\n'
                               "OP       DEDUPE(USED_RANGE($out), cols=[1], header=yes)\n"))
        self.assertIn("a spec may not both SET_FORMULA and DEDUPE: Excel rewrites formula references when rows move, "
                      "which a TEST can't prove (D43)", found)
        self.assertFalse([m for m in found if "a BUILD may not do both" in m])

    def test_a_wrong_sheet_test(self):   # rule 5
        text = changed("TEST     fixture=ncr_01.xlsx call=modMain.RunHeadless(Summary, NCR) assert=RESULT = -1 AND "
                       "ERROR_COUNT = 1\n", "")
        self.assertEqual(errors(text), ["modMain.RunHeadless has PRE and the fixture has another sheet, but no "
                                        "wrong-sheet TEST expects RESULT = -1 AND ERROR_COUNT = 1 (rule 5)"])


class Interpreter(unittest.TestCase):   # Phase 2: BUILD TESTs run through modMain.RunHeadless (D52)
    DAY = datetime.date(2026, 10, 21)

    def run_first(self, text):
        return run_tests(text, self.DAY)[0]

    def test_the_sample_proves_its_tests(self):
        results = run_tests(NCR_REVIEW, self.DAY)
        self.assertEqual([(r.passed, r.result, r.errors) for r in results], [(True, 0, 0), (True, 4, 0), (True, -1, 1)])

    def test_a_callee_error_is_logged_and_the_caller_goes_on(self):
        text = changed('OP       LET(cDays, COL($ncr, "Days Open"))', 'OP       LET(cDays, COL($ncr, "Days Open"))\nOP       LET(n, 1 / 0)')
        r = self.run_first(text)
        self.assertEqual((r.result, r.errors), (-2, 1))   # MarkOverdue gave -2 into marked; RunHeadless went on
        self.assertEqual(r.log[0][0], "MarkOverdue")
        self.assertIn("ASSERT: COUNTIF(NCR![Status], \"Overdue\") = 4 is false (left 0, right 4)", r.failed)
        # WriteSummary still ran: the header, Open and Closed, and no stale rows
        self.assertIn("ASSERT: COUNTA(Summary![Status]) = 4 is false (left 3, right 4)", r.failed)
        self.assertFalse(any("stale" in f for f in r.failed), r.failed)

    def test_a_callee_pre_failure_gives_minus_one(self):
        text = changed('PRE      COL($ncr, "Status") > 0; COL($ncr, "Days Open") > 0\nMODIFIES $ncr["Status"]\nOP       LET(cStatus',
                       'PRE      COL($ncr, "Owner") > 0\nMODIFIES $ncr["Status"]\nOP       LET(cStatus')
        r = self.run_first(text)
        self.assertEqual((r.result, r.errors), (-1, 1))
        self.assertEqual(r.log, [("MarkOverdue", 0, "PRE failed: condition 1")])

    def test_unchanged_follows_runheadless_modifies(self):
        text = changed('MODIFIES $ncr["Status"]\nMODIFIES $out\nOP       CALL', 'MODIFIES $out\nOP       CALL')
        r = self.run_first(text)
        self.assertFalse(r.passed)
        self.assertIn('UNCHANGED: NCR!B2 was "Open", now "Overdue"', r.failed)

    def test_values_convert_to_their_parameter_types(self):
        text = changed("CALL(modSteps.MarkOverdue, $ncr, 30, into=marked)", "CALL(modSteps.MarkOverdue, $ncr, 30.6, into=marked)")
        self.assertEqual(run_tests(text, self.DAY)[1].result, 3)   # 30.6 is 31 as a Long, so 31 days is not over

    def test_a_build_before_0_5_has_no_tests_to_run(self):
        with self.assertRaises(VclError) as caught:
            run_tests(NCR_REVIEW.replace("VCL 0.5", "VCL 0.4").replace(", into=marked", "")
                      .replace(", into=n", ""), self.DAY)
        self.assertIn("a BUILD runs its TESTs from VCL 0.5 on", str(caught.exception))


class Compiler(unittest.TestCase):   # Phase 3: one lint-clean module per BUILD, and mutation over it (D53)
    def test_one_module_with_every_procedure(self):
        name, module = compile_build(NCR_REVIEW, "ncr_review.vcl")
        self.assertEqual(name, "NcrReview")
        self.assertEqual(lint(module), [])
        for line in ('n = RunHeadless(ThisWorkbook.Worksheets("NCR Register"), ThisWorkbook.Worksheets("NCR Summary"))',
                     "marked = MarkOverdue(ncr, 30)", "Call WriteSummary(ncr, out)"):
            self.assertIn(line, module)
        for head in ("Public Sub Run()", "Public Function RunHeadless(", "Public Function MarkOverdue(",
                     "Public Sub WriteSummary("):
            self.assertEqual(module.count(head), 1, head)

    def test_what_waits_is_refused(self):
        for old, new, why in (("SIG      modMain.Run()", "FORM     frmMain CONTROLS btnRun:CommandButton\nSIG      modMain.Run()",
                               "FORM waits until after March 2027"),
                              ("TARGET   xlsm", "TARGET   xlam", "TARGET xlam waits")):
            with self.subTest(why):
                with self.assertRaises(CompileError) as caught:
                    compile_build(changed(old, new))
                self.assertIn(why, str(caught.exception))
        older = NCR_REVIEW.replace("VCL 0.5", "VCL 0.4").replace(", into=marked", "").replace(", into=n", "")
        with self.assertRaises(CompileError) as caught:
            compile_build(older)
        self.assertIn("a BUILD compiles from VCL 0.5 on", str(caught.exception))

    def test_vclc_builds_it_into_dist_builds(self):
        with tempfile.TemporaryDirectory() as tmp:
            spec = Path(tmp) / "ncr_review.vcl"
            spec.write_text(NCR_REVIEW, encoding="utf-8")
            out = Path(tmp) / "out"
            with contextlib.redirect_stdout(io.StringIO()):
                code = vclc.main(["vclc", "all", str(spec), "-o", str(out), "--today", "2026-10-21"])
            self.assertEqual(code, 0)
            self.assertTrue((out / "builds" / "NcrReview.bas").is_file())
            self.assertTrue((out / "expected" / "ncr_review.json").is_file())
            self.assertFalse((out / "forge_module.bas").exists())   # no TASK spec was given

    def test_the_tests_catch_every_mutant(self):   # Example 4 (D58): 34 of 34
        killed, total, survivors = mutate_report(NCR_REVIEW, snapshot=True, today=datetime.date(2026, 10, 21))
        self.assertEqual((killed, total, survivors), (34, 34, []))
        run_lines = {i + 1 for i in untested_lines(NCR_REVIEW)}
        self.assertEqual(len(run_lines), 2)   # modMain.Run's CALL and MSG: no TEST calls Run, so no mutant there


class ReviewGaps(unittest.TestCase):   # Phase 4 (D55): the gaps the item 2 reviews found (D47, D51)
    def test_row_of_and_dedupe_into_are_proven(self):
        self.assertEqual(problems(STOCK), [])
        results = run_tests(STOCK, datetime.date(2026, 10, 21))
        self.assertEqual([(r.passed, r.result) for r in results], [(True, 0), (True, 2), (True, 2), (True, -1)])
        self.assertEqual(mutate_report(STOCK, snapshot=True, today=datetime.date(2026, 10, 21))[:2], (8, 8))

    def test_they_compile(self):
        from vcl_codegen import compile_text
        vba = compile_text(STOCK)
        self.assertIn("removed = ForgeDedupe(ForgeUsedRange(ws), Array(cPart), True)", vba)
        self.assertIn("ws.Cells(hit.Row, cStatus)", vba)

    def test_they_need_0_5(self):
        older = STOCK.replace("VCL 0.5", "VCL 0.4", 1)
        found = errors(older)
        self.assertIn("ROW_OF is VCL 0.5: start the spec with 'VCL 0.5' to use it", found)
        self.assertIn("DEDUPE into= is VCL 0.5: start the spec with 'VCL 0.5' to use it", found)

    def test_a_header_the_spec_adds_needs_no_pre(self):   # D56
        self.assertEqual(problems(CHECKED), [])
        results = run_tests(CHECKED, datetime.date(2026, 10, 21))
        self.assertEqual([(r.passed, r.result) for r in results], [(True, 3), (True, 3), (True, -1)])
        self.assertEqual(mutate_report(CHECKED, snapshot=True, today=datetime.date(2026, 10, 21))[:2], (18, 18))

    def test_only_a_header_it_writes_to_that_sheet(self):
        found = errors(CHECKED.replace('SET_VALUE(CELL($ws, 1, cNew), "Checked")', 'SET_VALUE(CELL($ws, 1, cNew), "Done")'))
        self.assertIn('header "Checked" on $ws is used but not checked in PRE (rule 14)', found)
        found = errors(CHECKED.replace('SET_VALUE(CELL($ws, 1, cNew), "Checked")', 'SET_VALUE(CELL($ws, 2, cNew), "Checked")'))
        self.assertIn('header "Checked" on $ws is used but not checked in PRE (rule 14)', found)

    def test_its_column_is_not_there_to_declare(self):   # UNCHANGED finds columns by their header before the run
        found = errors(CHECKED.replace("MODIFIES $ws\n", 'MODIFIES $ws ["Checked"]\n'))
        self.assertEqual(found, ['MODIFIES $ws["Checked"]: the spec adds that header itself, so the column is not '
                                 'there before the run and no TEST can check it; declare the whole sheet with '
                                 'MODIFIES $ws (D56)'])

    def test_a_written_header_needs_0_5(self):
        found = errors(CHECKED.replace("VCL 0.5", "VCL 0.4", 1))
        self.assertEqual(found, ['header "Checked" on $ws is used but not checked in PRE (rule 14): a header the '
                                 'spec writes to row 1 itself needs VCL 0.5'])

    def test_a_sheet_name_with_spaces_in_quotes(self):   # D57
        self.assertEqual(problems(OPEN_NCRS), [])
        results = run_tests(OPEN_NCRS, datetime.date(2026, 10, 21))
        self.assertEqual([(r.passed, r.result) for r in results], [(True, 3), (True, 0), (True, -1)])
        self.assertEqual(mutate_report(OPEN_NCRS, snapshot=True, today=datetime.date(2026, 10, 21))[:2], (13, 13))

    def test_a_failed_check_names_the_sheet_as_written(self):
        text = OPEN_NCRS.replace("VALUE('Open NCRs'!B1) = 3 AND ROWS", "VALUE('Open NCRs'!B1) = 4 AND ROWS")
        r = run_tests(text, datetime.date(2026, 10, 21))[0]
        self.assertFalse(r.passed)
        self.assertIn("ASSERT: VALUE('Open NCRs'!B1) = 4 is false (left 3, right 4)", r.failed)

    def test_quoted_sheet_names_are_checked(self):
        cases = [(OPEN_NCRS.replace("COUNTA(Ncr!B:B) = PRE", "COUNTA('Ncr'!B:B) = PRE"),
                  "assertion: 'Ncr'!B:B: Ncr is one word, so write Ncr!B:B without quotes"),
                 (OPEN_NCRS.replace("ROWS_USED('Open NCRs')", "ROWS_USED('Open: NCRs')"),
                  "assertion: 'Open: NCRs': 'Open: NCRs' can't be a sheet name (S23)"),
                 (OPEN_NCRS.replace("ROWS_USED('Open NCRs') = 1", "'Open NCRs' = 1"),
                  "assertion: 'Open NCRs': a sheet name goes before ! or inside ROWS_USED, like 'Open NCRs'!A1"),
                 (OPEN_NCRS.replace('WRITE_ROW(out, 1, "Open NCRs", n)', "WRITE_ROW(out, 1, 'Open NCRs', n)"),
                  "'Open NCRs': text goes in \"double quotes\"; 'single quotes' are only for a sheet name in a "
                  "TEST assertion")]
        for text, message in cases:
            with self.subTest(message=message):
                self.assertIn(message, errors(text))

    def test_quoted_sheet_names_need_0_5(self):
        found = errors(OPEN_NCRS.replace("VCL 0.5", "VCL 0.4", 1))
        self.assertIn("assertion: 'Open NCRs'!A1: a sheet name in quotes is VCL 0.5: start the spec with 'VCL 0.5' "
                      "to use it", found)
        self.assertIn("assertion: 'Open NCRs': a sheet name in quotes is VCL 0.5: start the spec with 'VCL 0.5' "
                      "to use it", found)

    def test_row_of_nothing_stops_the_run(self):
        text = STOCK.replace("OP       IF NOT IS_NOTHING(hit)\nOP         SET_VALUE", "OP       IF TRUE\nOP         SET_VALUE")
        r = run_tests(text, datetime.date(2026, 10, 21))[2]   # P-999: FIND gives NOTHING
        self.assertEqual((r.result, r.errors), (-2, 1))
        self.assertIn("ROW_OF(NOTHING)", r.log[0][2])


class Run(unittest.TestCase):
    def test_run_only_starts_the_tool(self):
        found = errors(changed('OP       MSG("NCRs marked Overdue: " & n)',
                               'OP       MSG("NCRs marked Overdue: " & n)\nOP       LOG("done")'))
        self.assertIn("modMain.Run only CALLs modMain.RunHeadless, LETs, IFs and MSGs, so nothing in it goes "
                      "untested (D49); move LOG into RunHeadless", found)

    def test_run_passes_the_real_sheets_by_name(self):
        found = errors(changed('SHEET(THIS_BOOK, "NCR Summary")', "ACTIVE_SHEET"))
        self.assertIn('CALL modMain.RunHeadless: modMain.Run passes $out as SHEET(THIS_BOOK, "<real name>") (D49)',
                      found)


if __name__ == "__main__":
    unittest.main()
