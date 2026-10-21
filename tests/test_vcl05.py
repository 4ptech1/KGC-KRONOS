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
from vcl_interp import VclError, run_tests  # noqa: E402
from vba_lint import lint  # noqa: E402
from vcl_codegen import CompileError, compile_build  # noqa: E402
from vcl_mutate import mutate_report, untested_lines  # noqa: E402
import vclc  # noqa: E402

# A small tool: RunHeadless marks old open NCRs Overdue (modSteps.MarkOverdue) and rebuilds a
# summary of the statuses (modSteps.WriteSummary); modMain.Run starts it on the real sheets.
NCR_REVIEW = r"""VCL 0.5
BUILD
# ASSUME: NCR headers (row 1) include "NCR No", "Status" and "Days Open"
# ASSUME: an Open NCR older than the limit becomes Overdue; Summary is rebuilt on every run
NAME     NcrReview
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(ncr As Worksheet, out As Worksheet) As Long
SIG      modSteps.MarkOverdue(ncr As Worksheet, limit As Long) As Long
SIG      modSteps.WriteSummary(ncr As Worksheet, out As Worksheet)
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.MarkOverdue
FLOW     modMain.RunHeadless -> modSteps.WriteSummary
FIXTURE  ncr_01.xlsx sheet=NCR cols=A:NCR No,B:Status,C:Days Open rows=8 data=A:seq(1,1);B:values("Open"x5,"Closed"x2,""x1)+boundary("Open","Open","Closed","","Open");C:int(1,60)+boundary(31,30,45,10)
FIXTURE  ncr_01.xlsx sheet=Summary cols=A:Status,B:Count rows=0 stale=2
TEST     fixture=ncr_01.xlsx call=modMain.RunHeadless(NCR, Summary) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(NCR![Status], "Overdue") = 4 AND COUNTIF(NCR![Status], "Open") = 1 AND COUNTA(Summary![Status]) = 4 AND VALUE(Summary![Count]2) = 4 AND COUNTIF(Summary![Status], "stale") = 0
TEST     fixture=ncr_01.xlsx call=modMain.RunHeadless(NCR, Summary) assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(NCR![NCR No]) = PRE_COUNTA(NCR![NCR No]) AND VALUE(NCR![Status]2) = "Overdue" AND VALUE(NCR![Status]3) = "Open" AND VALUE(Summary![Status]2) = "Overdue"
TEST     fixture=ncr_01.xlsx call=modMain.RunHeadless(Summary, NCR) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.5
TASK
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "NCR Register"), SHEET(THIS_BOOK, "NCR Summary"), into=n)
OP       MSG("NCRs marked Overdue: " & n)
ERR      standard
---
VCL 0.5
TASK
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     ncr As Worksheet, out As Worksheet
RET      Long
VAR      marked As Long
PRE      COL($ncr, "Status") > 0; COL($ncr, "Days Open") > 0
MODIFIES $ncr["Status"]
MODIFIES $out
OP       CALL(modSteps.MarkOverdue, $ncr, 30, into=marked)
OP       CALL(modSteps.WriteSummary, $ncr, $out)
OP       RETURN(marked)
ERR      standard
---
VCL 0.5
TASK
PROC     MarkOverdue
KIND     Function
MODULE   modSteps
ARGS     ncr As Worksheet, limit As Long
RET      Long
VAR      cStatus As Long
VAR      cDays As Long
VAR      n As Long
PRE      COL($ncr, "Status") > 0; COL($ncr, "Days Open") > 0
MODIFIES $ncr["Status"]
OP       LET(cStatus, COL($ncr, "Status"))
OP       LET(cDays, COL($ncr, "Days Open"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, cStatus) var=r
OP         IF SAME(CELL($ncr, r, cStatus), "Open")
OP           IF IS_NUMBER(CELL($ncr, r, cDays))
OP             IF CELL($ncr, r, cDays) > $limit
OP               SET_VALUE(CELL($ncr, r, cStatus), "Overdue")
OP               LET(n, n + 1)
OP       RETURN(n)
ERR      standard
---
VCL 0.5
TASK
PROC     WriteSummary
KIND     Sub
MODULE   modSteps
ARGS     ncr As Worksheet, out As Worksheet
VAR      cStatus As Long
VAR      counts As Dictionary
VAR      got As Long
VAR      outRow As Long
PRE      COL($ncr, "Status") > 0
MODIFIES $out
OP       LET(cStatus, COL($ncr, "Status"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Status", "Count")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, cStatus) var=r
OP         SKIP_IF IS_EMPTY(CELL($ncr, r, cStatus))
OP         DICT_ADD(counts, TRIM(CELL($ncr, r, cStatus)), 1, mode=count)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         WRITE_ROW($out, outRow, k, got)
OP         LET(outRow, outRow + 1)
ERR      standard
"""


# Phase 4a (D54): ROW_OF finds the row of a FIND result; DEDUPE into= counts the rows it removed.
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
        paths = sorted((ROOT / "specs").glob("*.vcl")) + [ROOT / "vcl" / "examples" / name
                                                         for name in ("example1_rev3.vcl", "example3_files.vcl")]
        day = datetime.date(2026, 10, 21)
        for path in paths:
            first, rest = path.read_text(encoding="utf-8").split("\n", 1)
            with self.subTest(spec=path.name):
                newer = "VCL 0.5\n" + rest
                text = first + "\n" + rest
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

    def test_the_tests_catch_every_mutant_that_can_be_caught(self):
        killed, total, survivors = mutate_report(NCR_REVIEW, snapshot=True, today=datetime.date(2026, 10, 21))
        self.assertEqual((killed, total), (total - 1, total))
        line = NCR_REVIEW.splitlines()[survivors[0][0] - 1]
        self.assertIn("mode=count", line)   # the one survivor: DICT_ADD's value, which mode=count ignores
        run_lines = {i + 1 for i in untested_lines(NCR_REVIEW)}
        self.assertEqual(len(run_lines), 2)   # modMain.Run's CALL and MSG: no TEST calls Run, so no mutant there


class ReviewGaps(unittest.TestCase):   # Phase 4 (D54): the gaps the item 2 reviews found (D47, D51)
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
