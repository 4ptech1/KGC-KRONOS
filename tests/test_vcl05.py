"""VEP 0.5 (docs/vep-0.5-build.md): CALL and BUILD specs. Phase 1: the VCL 0.5 line, CALL into=,
and the validator and type-checker rules that let a BUILD's TESTs prove the whole tool (D49)."""
import datetime
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
import validator  # noqa: E402
from vcl_interp import run_tests  # noqa: E402

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
SIG      modSteps.WriteSummary(ncr As Worksheet, out As Worksheet) As Long
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.MarkOverdue
FLOW     modMain.RunHeadless -> modSteps.WriteSummary
FIXTURE  ncr_01.xlsx sheet=NCR cols=A:NCR No,B:Status,C:Days Open rows=8 data=A:seq(1,1);B:values("Open"x5,"Closed"x2,""x1)+boundary("Open","Open","Closed","","Open");C:int(1,60)+boundary(30,31,45,10)
FIXTURE  ncr_01.xlsx sheet=Summary cols=A:Status,B:Count rows=0 stale=2
TEST     fixture=ncr_01.xlsx call=modMain.RunHeadless(NCR, Summary) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(Summary![Status], "stale") = 0
TEST     fixture=ncr_01.xlsx call=modMain.RunHeadless(NCR, Summary) assert=COUNTA(NCR![NCR No]) = PRE_COUNTA(NCR![NCR No]) AND RESULT >= 1
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
VAR      written As Long
PRE      COL($ncr, "Status") > 0; COL($ncr, "Days Open") > 0
MODIFIES $ncr["Status"]
MODIFIES $out
OP       CALL(modSteps.MarkOverdue, $ncr, 30, into=marked)
OP       CALL(modSteps.WriteSummary, $ncr, $out, into=written)
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
KIND     Function
MODULE   modSteps
ARGS     ncr As Worksheet, out As Worksheet
RET      Long
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
OP       RETURN(outRow - 2)
ERR      standard
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
        found = errors(changed("CALL(modSteps.WriteSummary, $ncr, $out, into=written)",
                               'CALL(modSteps.WriteSummary, $ncr, SHEET(THIS_BOOK, "Summary"), into=written)'))
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
