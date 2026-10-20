"""VEP 0.4 (docs/vep-0.4-rev3-rest.md). Phase 1: the interpreter runs what the compiler already
compiled: INSERT_ROW, READ_ARRAY, WRITE_ARRAY, ARR, ROWS, COLS, SELECT_CASE, LOG, MSG, RANGE, SHEET,
FILE_NAME and FOR_EACH over SHEETS and CELLS (D36, D37). Phase 2: the VCL 0.4 line, lists in
DEDUPE cols=[...] and DEDUPE itself (S21, D38). Phase 3: FIND, IS_NOTHING and NOTHING (S22, D39). Phase 4a: Workbook
arguments, MODIFIES $book sheets, ADD_SHEET and SHEET_EXISTS (S23, D40). Phase 4b: COPY_SHEET (D41). Phase 5:
FILTER, hidden rows and VISIBLE_COUNT (D42)."""
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
import validator  # noqa: E402
from vba_lint import lint  # noqa: E402
from vcl_codegen import assemble, compile_text  # noqa: E402
from vcl_interp import modifies_of, needs_unchanged_check, run_tests  # noqa: E402
from vcl_parse import parse  # noqa: E402
from vcl_mutate import mutate_report  # noqa: E402
from vcl_types import check_text  # noqa: E402

FORGELIB = (ROOT / "vba" / "ForgeLib.bas").read_text(encoding="utf-8")

# Arrays, SELECT_CASE, INSERT_ROW and LOG: grade each weld from its Result, on a copy in Report.
GRADE = r"""VCL 0.3
TASK
# ASSUME: Welds headers (row 1) include "Result" and "Grade"
# ASSUME: Result is matched as typed, after TRIM; anything not listed gets "Unknown"
PROC     GradeWelds
KIND     Function
MODULE   modGrades
ARGS     ws As Worksheet, out As Worksheet
RET      Long
VAR      cResult As Long
VAR      cGrade As Long
VAR      data As Variant
VAR      n As Long
PRE      COL($ws, "Result") > 0; COL($ws, "Grade") > 0
MODIFIES $out
OP       LET(cResult, COL($ws, "Result"))
OP       LET(cGrade, COL($ws, "Grade"))
OP       CLEAR(USED_RANGE($out))
OP       READ_ARRAY(USED_RANGE($ws), into=data)
OP       LOOP_ROWS dir=down from=2 to=ROWS(data) var=r
OP         SELECT_CASE TRIM(ARR(data, r, cResult))
OP           CASE "Accepted"
OP             SET_VALUE(ARR(data, r, cGrade), "A")
OP           CASE "Rejected", "Repair"
OP             SET_VALUE(ARR(data, r, cGrade), "R-1")
OP             LET(n, n + 1)
OP           CASE_ELSE
OP             SET_VALUE(ARR(data, r, cGrade), "Unknown")
OP       WRITE_ARRAY(data, to=CELL($out, 1, 1))
OP       INSERT_ROW($out, 1)
OP       SET_VALUE(CELL($out, 1, 1), "Graded")
OP       LOG("graded " & n)
OP       RETURN(n)
ERR      standard
FIXTURE  grades_01.xlsx sheet=Welds cols=A:Seam,B:Result,C:Grade rows=8 data=A:seq(1,1);B:values("Accepted"x3,"Rejected"x2,"Repair"x1,"accepted"x1,""x1)+boundary("Accepted","Rejected","Repair","accepted","")
FIXTURE  grades_01.xlsx sheet=Report cols=A:Old rows=0 stale=3
TEST     fixture=grades_01.xlsx call=GradeWelds(Welds, Report) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Report!C:C, "R-1") = 3 AND COUNTIF(Report!C:C, "A") = 3 AND COUNTIF(Report!C:C, "Unknown") = 2 AND VALUE(Report!A1) = "Graded" AND VALUE(Report!A2) = "Seam"
TEST     fixture=grades_01.xlsx call=GradeWelds(Welds, Report) assert=COUNTA(Welds![Result]) = PRE_COUNTA(Welds![Result]) AND COUNTA(Welds![Grade]) = PRE_COUNTA(Welds![Grade])
TEST     fixture=grades_01.xlsx call=GradeWelds(Report, Report) assert=RESULT = -1 AND ERROR_COUNT = 1
"""

# FOR_EACH over SHEETS and CELLS, RANGE, SHEET, FILE_NAME and MSG.
MARKS = r"""VCL 0.3
TASK
# ASSUME: a sheet "has marks" when its row 1 includes the header "Mark"
# ASSUME: the marks to count are in Checks!B2:B6, matched with SAME; D2 gets the file name of the path in C2
PROC     CountMarks
KIND     Function
MODULE   modMarks
ARGS     ws As Worksheet
RET      Long
VAR      sheetsWithMarks As Long
VAR      ticks As Long
VAR      notes As Long
PRE      COL($ws, "Mark") > 0
MODIFIES $ws
OP       FOR_EACH var=sh in=SHEETS(THIS_BOOK)
OP         IF COL(sh, "Mark") > 0 THEN LET(sheetsWithMarks, sheetsWithMarks + 1)
OP       FOR_EACH var=v in=CELLS(RANGE($ws, "B2:B6"))
OP         IF SAME(v, "x") THEN LET(ticks, ticks + 1)
OP       LET(notes, LASTROW(SHEET(THIS_BOOK, "Notes"), 1))
OP       SET_VALUE(RANGE($ws, "D2"), FILE_NAME(TRIM(CELL($ws, 2, 3))))
OP       MSG("ticks: " & ticks)
OP       RETURN(sheetsWithMarks * 100 + ticks * 10 + notes)
ERR      standard
FIXTURE  marks_01.xlsx sheet=Checks cols=A:Item,B:Mark,C:Path,D:File rows=6 data=A:seq(1,1);C:values("C:\Docs\WPS-7.pdf"x1,"D:\a\b.pdf"x5)+boundary("C:\Docs\WPS-7.pdf");B:values("x"x2," X "x1,"o"x2,""x1)+boundary("x"," X ","o","x","o","")
FIXTURE  marks_01.xlsx sheet=Notes cols=A:Note rows=3 data=A:text(5)
FIXTURE  marks_01.xlsx sheet=Spare cols=A:Mark rows=2 data=A:values("x"x2)
TEST     fixture=marks_01.xlsx call=CountMarks(Checks) runs=2 assert=RESULT = 234 AND ERROR_COUNT = 0 AND VALUE(Checks!D2) = "WPS-7.pdf"
TEST     fixture=marks_01.xlsx call=CountMarks(Checks) assert=COUNTA(Checks![Mark]) = PRE_COUNTA(Checks![Mark]) AND COUNTA(Checks![File]) = PRE_COUNTA(Checks![File]) + 1
TEST     fixture=marks_01.xlsx call=CountMarks(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
"""

# DEDUPE with a list of key columns (VCL 0.4): drop repeated welds, never a row with a blank key.
DEDUPE = r"""VCL 0.4
TASK
# ASSUME: a weld is listed twice when Seam and Welder repeat (case ignored, as typed); the first listing stays
# ASSUME: a row with a blank Seam or Welder is never removed
PROC     DropRepeatedWelds
KIND     Function
MODULE   modDedupe
ARGS     ws As Worksheet
RET      Long
VAR      before As Long
PRE      COL($ws, "Seam") > 0; COL($ws, "Welder") > 0
MODIFIES $ws
OP       LET(before, LASTROW($ws, 3))
OP       DEDUPE(USED_RANGE($ws), cols=[COL($ws, "Seam"), COL($ws, "Welder")], header=yes)
OP       RETURN(before - LASTROW($ws, 3))
ERR      standard
FIXTURE  welds_01.xlsx sheet=Welds cols=A:Seam,B:Welder,C:Note rows=8 data=C:seq(1,1);A:values("S1"x2,"s1"x1,"S2"x1,"S1 "x1,""x2,"S3"x1)+boundary("S1","S1","s1","S2","S1 ","","","S3");B:values("W1"x7,""x1)+boundary("W1","W1","W1","W1","W1","W1","W1","")
FIXTURE  welds_01.xlsx sheet=Other cols=A:Item rows=2 data=A:seq(1,1)
TEST     fixture=welds_01.xlsx call=DropRepeatedWelds(Welds) assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(Welds!C2) = 1 AND VALUE(Welds!C3) = 4 AND VALUE(Welds!A4) = "S1 " AND VALUE(Welds!C7) = 8
TEST     fixture=welds_01.xlsx call=DropRepeatedWelds(Welds) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTA(Welds!C:C) = PRE_COUNTA(Welds!C:C) - 2
TEST     fixture=welds_01.xlsx call=DropRepeatedWelds(Other) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Other!A:A) = PRE_COUNTA(Other!A:A)
"""

# FIND written in place, so MODIFIES can tell the sheet: highlight the first matching NCR.
FIND = r"""VCL 0.3
TASK
# ASSUME: NCR numbers are in A2:A9 of the register and notes in B2:B9; the number is matched with SAME
# ASSUME: the first matching NCR is highlighted yellow; a number that is not there gives 0 and changes nothing
PROC     FlagNcr
KIND     Function
MODULE   modNcr
ARGS     ws As Worksheet, ncrNo As String
RET      Long
VAR      n As Long
PRE      COL($ws, "NCR No") > 0
MODIFIES $ws fills
OP       IF IS_NOTHING(FIND(RANGE($ws, "A2:A9"), $ncrNo, match=whole)) THEN RETURN(0)
OP       FORMAT(FIND(RANGE($ws, "A2:A9"), $ncrNo, match=whole), fill=yellow)
OP       LET(n, 1)
OP       IF NOT IS_NOTHING(FIND(RANGE($ws, "B2:B9"), $ncrNo, match=part)) THEN LET(n, 2)
OP       RETURN(n)
ERR      standard
FIXTURE  ncr_01.xlsx sheet=Ncr cols=A:NCR No,B:Note rows=8 data=A:values("NCR-001"x1,"NCR-0071"x1," NCR-007 "x1,"NCR-007"x1,"NCR-010"x1,""x1,"NCR-012"x1,"NCR-013"x1)+boundary("NCR-001","NCR-0071"," NCR-007 ","NCR-007","NCR-010","","NCR-012","NCR-013");B:values("ok"x5,"Rework per ncr-007"x1,""x2)+boundary("ok","ok","ok","ok","ok","Rework per ncr-007","","")
FIXTURE  ncr_01.xlsx sheet=Other cols=A:Item rows=2 data=A:seq(1,1)
TEST     fixture=ncr_01.xlsx call=FlagNcr(Ncr, "ncr-007") runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND FILL_COUNT(Ncr!A4, #FFFF00) = 1 AND FILL_COUNT(Ncr!A2:A9, #FFFF00) = 1
TEST     fixture=ncr_01.xlsx call=FlagNcr(Ncr, "NCR-010") assert=RESULT = 1 AND ERROR_COUNT = 0 AND FILL_COUNT(Ncr!A6, #FFFF00) = 1 AND COUNTA(Ncr!A:A) = PRE_COUNTA(Ncr!A:A)
TEST     fixture=ncr_01.xlsx call=FlagNcr(Ncr, "NCR-099") assert=RESULT = 0 AND ERROR_COUNT = 0 AND FILL_COUNT(Ncr!A2:B9, #FFFF00) = 0
TEST     fixture=ncr_01.xlsx call=FlagNcr(Other, "NCR-001") assert=RESULT = -1 AND ERROR_COUNT = 1
"""

# ADD_SHEET (VCL 0.4): a new Summary sheet, unless the workbook has one already.
ADD_SHEET = r"""VCL 0.4
TASK
# ASSUME: the count goes to a new sheet "Summary" after the register; an existing Summary is left as it is
# ASSUME: an NCR is open when its Status is SAME as "Open"
PROC     AddNcrSummary
KIND     Function
MODULE   modSummary
ARGS     book As Workbook, ncr As Worksheet
RET      Long
VAR      out As Worksheet
VAR      n As Long
PRE      COL($ncr, "Status") > 0
MODIFIES $book sheets
OP       IF SHEET_EXISTS($book, "Summary") THEN RETURN(0)
OP       ADD_SHEET($book, name="Summary", after=$ncr, into=out)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, COL($ncr, "Status")) var=r
OP         IF SAME(CELL($ncr, r, COL($ncr, "Status")), "Open") THEN LET(n, n + 1)
OP       WRITE_ROW(out, 1, "Open NCRs", n)
OP       RETURN(n)
ERR      standard
FIXTURE  ncr_01.xlsx sheet=Ncr cols=A:NCR No,B:Status rows=6 data=A:seq(1,1);B:values("Open"x2," open "x1,"Closed"x2,""x1)+boundary("Open")
FIXTURE  ncr_01.xlsx sheet=Other cols=A:Item rows=2 data=A:seq(1,1)
FIXTURE  ncr_02.xlsx sheet=Ncr cols=A:NCR No,B:Status rows=2 data=A:seq(1,1);B:values("Open"x2)
FIXTURE  ncr_02.xlsx sheet=Summary cols=A:Old rows=1 data=A:values("keep"x1)
TEST     fixture=ncr_01.xlsx call=AddNcrSummary(ncr_01.xlsx, Ncr) assert=RESULT = 3 AND ERROR_COUNT = 0 AND SHEET_EXISTS("Summary") AND VALUE(Summary!A1) = "Open NCRs" AND VALUE(Summary!B1) = 3 AND COUNTA(Ncr!B:B) = PRE_COUNTA(Ncr!B:B)
TEST     fixture=ncr_01.xlsx call=AddNcrSummary(ncr_01.xlsx, Ncr) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE(Summary!B1) = 3
TEST     fixture=ncr_02.xlsx call=AddNcrSummary(ncr_02.xlsx, Ncr) assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE(Summary!A2) = "keep"
TEST     fixture=ncr_01.xlsx call=AddNcrSummary(ncr_01.xlsx, Other) assert=RESULT = -1 AND ERROR_COUNT = 1
"""

# COPY_SHEET (VCL 0.4): a job sheet copied from Template, named by the job number.
COPY_SHEET = r"""VCL 0.4
TASK
# ASSUME: a job sheet is a copy of Template named by the job number (TRIM), placed after Template
# ASSUME: an existing job sheet is left as it is; the copy gets "Job" in C1 and the job number in C2
PROC     NewJobSheet
KIND     Function
MODULE   modJobs
ARGS     book As Workbook, tpl As Worksheet, jobNo As String
RET      Long
VAR      job As Worksheet
VAR      title As String
PRE      COL($tpl, "Item") > 0
MODIFIES $book sheets
OP       LET(title, TRIM($jobNo))
OP       IF SHEET_EXISTS($book, title) THEN RETURN(0)
OP       COPY_SHEET($tpl, name=title, after=$tpl, into=job)
OP       SET_VALUE(CELL(job, 1, 3), "Job")
OP       SET_VALUE(CELL(job, 2, 3), title)
OP       RETURN(LASTROW(job, 1))
ERR      standard
FIXTURE  job_01.xlsx sheet=Template cols=A:Item,B:Qty rows=3 data=A:values("Bolt"x1,"Nut"x1,"Pipe"x1)+boundary("Bolt","Nut","Pipe");B:values("5"x2,""x1)+boundary("5","5","")
FIXTURE  job_01.xlsx sheet=Other cols=A:Note rows=1 data=A:values("x"x1)
FIXTURE  job_02.xlsx sheet=Template cols=A:Item rows=1 data=A:values("Bolt"x1)
FIXTURE  job_02.xlsx sheet=J101 cols=A:Old rows=1 data=A:values("keep"x1)
TEST     fixture=job_01.xlsx call=NewJobSheet(job_01.xlsx, Template, " J101 ") assert=RESULT = 4 AND ERROR_COUNT = 0 AND SHEET_EXISTS("J101") AND VALUE(J101!A2) = "Bolt" AND VALUE(J101!A4) = "Pipe" AND VALUE(J101!C1) = "Job" AND VALUE(J101!C2) = "J101" AND COUNTA(Template!C:C) = PRE_COUNTA(Template!C:C)
TEST     fixture=job_01.xlsx call=NewJobSheet(job_01.xlsx, Template, "J101") runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE(J101!C2) = "J101"
TEST     fixture=job_02.xlsx call=NewJobSheet(job_02.xlsx, Template, "j101") assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE(J101!A2) = "keep"
TEST     fixture=job_01.xlsx call=NewJobSheet(job_01.xlsx, Template, "J/101") assert=RESULT = -2 AND ERROR_COUNT = 1
TEST     fixture=job_01.xlsx call=NewJobSheet(job_01.xlsx, Other, "J102") assert=RESULT = -1 AND ERROR_COUNT = 1
"""

# FILTER (D42): only one welder's welds stay visible; nothing is deleted.
FILTER = r"""VCL 0.4
TASK
# ASSUME: only the welds of one welder stay visible, matched with SAME on the Welder column
# ASSUME: nothing is deleted and the header row stays visible; the result is the number of visible welds
PROC     ShowWelderWelds
KIND     Function
MODULE   modFilter
ARGS     ws As Worksheet, welder As String
RET      Long
VAR      cWelder As Long
VAR      n As Long
PRE      COL($ws, "Welder") > 0
MODIFIES $ws
OP       LET(cWelder, COL($ws, "Welder"))
OP       FILTER(USED_RANGE($ws), col=cWelder, crit=$welder)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cWelder) var=r
OP         IF SAME(CELL($ws, r, cWelder), $welder) THEN LET(n, n + 1)
OP       RETURN(n)
ERR      standard
FIXTURE  welds_01.xlsx sheet=Welds cols=A:Seam,B:Welder rows=6 data=A:seq(1,1);B:values("W1"x2," w1"x1,"W2"x2,""x1)+boundary("W1","W2"," w1","W2","","W1")
FIXTURE  welds_01.xlsx sheet=Other cols=A:Item rows=1 data=A:seq(1,1)
TEST     fixture=welds_01.xlsx call=ShowWelderWelds(Welds, " w1 ") assert=RESULT = 3 AND ERROR_COUNT = 0 AND VISIBLE_COUNT(Welds!A:A) = 4 AND COUNTA(Welds!A:A) = PRE_COUNTA(Welds!A:A)
TEST     fixture=welds_01.xlsx call=ShowWelderWelds(Welds, "W2") runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND VISIBLE_COUNT(Welds!B:B) = 3
TEST     fixture=welds_01.xlsx call=ShowWelderWelds(Welds, "W9") assert=RESULT = 0 AND ERROR_COUNT = 0 AND VISIBLE_COUNT(Welds!A:A) = 1
TEST     fixture=welds_01.xlsx call=ShowWelderWelds(Other, "W1") assert=RESULT = -1 AND ERROR_COUNT = 1
"""

FIXTURE = 'FIXTURE  t_01.xlsx sheet=Data cols=A:Key,B:Qty rows=4 data=A:values("a"x2,"b"x2);B:int(1,9)'


def ops(*lines, head=(), version="VCL 0.3"):
    """A small spec around OP lines, with one TEST, to watch the interpreter."""
    return "\n".join([version, "TASK", "PROC     P", "KIND     Function", "MODULE   modP",
                      "ARGS     ws As Worksheet", "RET      Long", "VAR      n As Long", *head, "MODIFIES $ws"]
                     + ["OP       " + o for o in lines]
                     + ["OP       RETURN(n)", "ERR      standard", FIXTURE,
                        "TEST     fixture=t_01.xlsx call=P(Data) assert=ERROR_COUNT = 0"])


def run(text):
    (result,) = run_tests(text)
    return result


def cells(result, sheet="Data"):
    return dict(result.state[sheet][0])


class Samples(unittest.TestCase):
    def test_they_validate_pass_and_catch_every_mutant(self):
        for name, text in (("GRADE", GRADE), ("MARKS", MARKS), ("DEDUPE", DEDUPE), ("FIND", FIND),
                           ("ADD_SHEET", ADD_SHEET), ("COPY_SHEET", COPY_SHEET), ("FILTER", FILTER)):
            with self.subTest(name):
                self.assertEqual([p.message for p in validator.validate_text(text)], [])
                results = run_tests(text)
                self.assertTrue(all(r.passed for r in results), [(r.line, r.failed) for r in results])
                killed, total, survivors = mutate_report(text, snapshot=True)
                self.assertEqual((killed, survivors), (total, []))

    def test_they_compile_lint_clean(self):
        vba = compile_text(GRADE) + compile_text(MARKS)
        self.assertEqual(lint(assemble([compile_text(GRADE), compile_text(MARKS)])), [])
        self.assertIn("data = ForgeReadArray(ForgeUsedRange(ws))", vba)
        self.assertIn("forgeSel1 = ForgeTrim(data(r, cResult))", vba)     # the selector, evaluated once
        self.assertIn("Dim forgeSel1 As Variant", vba)
        self.assertIn("Select Case True", vba)
        self.assertIn('Case (forgeSel1 = "Rejected"), (forgeSel1 = "Repair")', vba)
        self.assertIn("Case Else", vba)
        self.assertIn("ForgeInsertRow out, 1", vba)
        self.assertIn("ForgeWriteArray data, out.Cells(1, 1)", vba)
        self.assertIn("For Each sh In ThisWorkbook.Worksheets", vba)
        self.assertIn('For Each forgeCell_v In ws.Range("B2:B6").Cells', vba)
        self.assertIn("v = forgeCell_v.Value", vba)                       # the loop variable is the value
        self.assertIn('ForgeLastRow(ThisWorkbook.Worksheets("Notes"), 1)', vba)
        self.assertIn("ForgeFileName(ForgeTrim(ws.Cells(2, 3).Value))", vba)
        vba = compile_text(DEDUPE)
        self.assertEqual(lint(assemble([vba])), [])
        self.assertIn('ForgeDedupe ForgeUsedRange(ws), Array(ForgeCol(ws, "Seam"), ForgeCol(ws, "Welder")), True', vba)
        vba = compile_text(FIND)
        self.assertEqual(lint(assemble([vba])), [])
        self.assertIn('If (ForgeFind(ws.Range("A2:A9"), ncrNo, True) Is Nothing) Then', vba)
        self.assertIn('If (Not (ForgeFind(ws.Range("B2:B9"), ncrNo, False) Is Nothing)) Then n = 2', vba)
        vba = compile_text(ADD_SHEET)
        self.assertEqual(lint(assemble([vba])), [])
        self.assertIn("Public Function AddNcrSummary(ByVal book As Workbook, ByVal ncr As Worksheet) As Long", vba)
        self.assertIn('If ForgeSheetExists(book, "Summary") Then', vba)
        self.assertIn('Set out = ForgeAddSheet(book, "Summary", ncr)', vba)
        vba = compile_text(COPY_SHEET)
        self.assertEqual(lint(assemble([vba])), [])
        self.assertIn("Set job = ForgeCopySheet(tpl, title, tpl)", vba)
        vba = compile_text(FILTER)
        self.assertEqual(lint(assemble([vba])), [])
        self.assertIn("ForgeFilter ForgeUsedRange(ws), cWelder, welder", vba)

    def test_log_and_msg_are_never_errors(self):
        grade, marks = run_tests(GRADE)[0], run_tests(MARKS)[0]
        self.assertEqual((grade.errors, grade.messages), (0, [("LOG", "graded 3")]))
        self.assertEqual((marks.errors, marks.messages), (0, [("MSG", "ticks: 3")]))


class Interpreter(unittest.TestCase):
    def test_insert_row_is_blank_and_moves_fills(self):
        r = run(ops("FORMAT(ROW_RANGE($ws, 2, 1, 2), fill=red)", "INSERT_ROW($ws, 2)"))
        got = cells(r)
        self.assertEqual((got[(1, 1)], got.get((2, 1)), got[(3, 1)]), ("Key", None, "a"))
        self.assertEqual(sorted(dict(r.state["Data"][1])), [(3, 1), (3, 2)])   # the red row moved; the new one has none

    def test_read_array_is_2d_even_for_one_cell(self):
        text = ops("READ_ARRAY(CELL($ws, 1, 1), into=a)", "LET(n, ROWS(a) * 10 + COLS(a))", head=["VAR      a As Variant"])
        self.assertEqual(run(text).result, 11)

    def test_arr_outside_the_array_is_an_error(self):
        text = ops("READ_ARRAY(CELL($ws, 1, 1), into=a)", "SET_VALUE(ARR(a, 2, 1), 0)", head=["VAR      a As Variant"])
        r = run(text)
        self.assertEqual((r.result, r.errors), (-2, 1))

    def test_write_array_keeps_text_as_text(self):
        text = ops("READ_ARRAY(CELL($ws, 1, 1), into=a)", 'SET_VALUE(ARR(a, 1, 1), "1001")',
                   "WRITE_ARRAY(a, to=CELL($ws, 1, 3))", head=["VAR      a As Variant"])
        self.assertEqual(cells(run(text))[(1, 3)], "1001")
        body = FORGELIB[FORGELIB.index("Public Sub ForgeWriteArray("):]
        self.assertIn("If ForgeNeedsTextGuard(out(r, c)) Then out(r, c) = \"'\" & out(r, c)", body[:body.index("End Sub")])

    def test_select_case_compares_as_equals(self):
        def chosen(selector, *cases):
            lines = ["SELECT_CASE " + selector]
            for i, case in enumerate(cases, start=1):
                lines += ["  " + case, f"    LET(n, {i})"]
            return run(ops(*lines)).result
        self.assertEqual(chosen('"25"', 'CASE "25.0"', "CASE_ELSE"), 2)        # text, not numbers (S7)
        self.assertEqual(chosen('"b"', 'CASE "a", "b"', 'CASE "b"'), 1)        # the first match wins
        self.assertEqual(chosen('"B"', 'CASE "b"'), 0)                         # = is case-sensitive; no CASE_ELSE
        self.assertEqual(chosen("2 + 3", "CASE 4, 5", "CASE_ELSE"), 1)

    def test_select_case_values_are_typed_as_equals(self):
        found = check_text(ops("SELECT_CASE CELL($ws, 2, 2)", "  CASE 5", "    LET(n, 1)"))
        self.assertTrue(any("raw cell value" in m for _, m in found), found)
        self.assertEqual(check_text(ops("SELECT_CASE CELL($ws, 2, 1)", '  CASE "a"', "    LET(n, 1)")), [])

    def test_ranges_inside_the_sheet_only(self):
        self.assertEqual(cells(run(ops('SET_VALUE(RANGE($ws, "C2:$C$3"), 7)')))[(3, 3)], "7")
        for ref in ("A:A", "1:1", "XFE1", "A1048577", "Totals"):
            with self.subTest(ref):
                r = run(ops(f'SET_VALUE(RANGE($ws, "{ref}"), 7)'))
                self.assertEqual((r.result, r.errors), (-2, 1))

    def test_books_and_sheets(self):
        self.assertEqual(run(ops('LET(n, LASTROW(SHEET(THIS_BOOK, "data"), 1))')).result, 5)   # name, any case
        self.assertEqual(run(ops('LET(n, LASTROW(SHEET(THIS_BOOK, "Missing"), 1))')).result, -2)
        self.assertEqual(run(ops("LET(n, LASTROW(ACTIVE_SHEET, 1))")).result, -2)

    def test_a_sheet_loop_variable_is_a_sheet(self):
        text = ops("FOR_EACH var=sh in=SHEETS(THIS_BOOK)", '  IF COL(sh, "Qty") > 0 THEN LET(n, n + 1)')
        self.assertEqual(check_text(text), [])
        self.assertEqual(run(text).result, 1)


def errors(text):
    """The errors on the OP lines (the small specs of ops() break the spec-wide TEST rules on purpose)."""
    first_op = text.splitlines().index(next(line for line in text.splitlines() if line.startswith("OP "))) + 1
    return [p.message for p in validator.validate_text(text) if p.severity == "error" and p.line >= first_op
            and "TEST" not in p.message]


class Dedupe(unittest.TestCase):
    KEYS = 'FIXTURE  t_01.xlsx sheet=Data cols=A:Key,B:Qty rows=6 data=A:values("a"x2,"A"x1,"b"x2,""x1)+boundary("a","A","b","a","","b");B:values("1"x5,"2"x1)+boundary("1","1","1","2","1","1")'

    def run_keys(self, op, version="VCL 0.4"):
        text = ops(op, head=[], version=version).replace(FIXTURE, self.KEYS)
        return run(text)

    def test_the_first_row_of_each_key_stays(self):
        r = self.run_keys("DEDUPE(USED_RANGE($ws), cols=[1, 2], header=yes)")
        self.assertEqual(r.errors, 0)
        rows = [(cells(r).get((i, 1)), cells(r).get((i, 2))) for i in range(1, 7)]
        self.assertEqual(rows, [("Key", "Qty"), ("a", "1"), ("b", "1"), ("a", "2"), (None, "1"), (None, None)])

    def test_one_key_column_and_no_header(self):
        r = self.run_keys("DEDUPE(USED_RANGE($ws), cols=1, header=no)")
        self.assertEqual([cells(r).get((i, 1)) for i in range(1, 6)], ["Key", "a", "b", None, None])

    def test_the_old_form_means_the_same_before_0_4(self):
        old = self.run_keys("DEDUPE(USED_RANGE($ws), cols=1,2, header=yes)", version="VCL 0.3")
        new = self.run_keys("DEDUPE(USED_RANGE($ws), cols=[1, 2], header=yes)")
        self.assertEqual(old.state, new.state)

    def test_a_key_column_beyond_the_sheet_is_an_error(self):
        for cols in ("[1, 16385]", "[0, 1]"):
            with self.subTest(cols):
                r = self.run_keys(f"DEDUPE(USED_RANGE($ws), cols={cols}, header=yes)")
                self.assertEqual((r.result, r.errors), (-2, 1))
        r = self.run_keys("DEDUPE(CELL($ws, 1, 1), cols=[0], header=yes)")   # no row to check: nothing read
        self.assertEqual(r.errors, 0)

    def test_the_validator_and_the_version_line(self):
        dedupe = "DEDUPE(USED_RANGE($ws), {}, header=yes)"
        self.assertIn("DEDUPE cols=[...] is VCL 0.4: start the spec with 'VCL 0.4' to use it",
                      errors(ops(dedupe.format("cols=[1, 2]"))))
        self.assertIn("DEDUPE cols=1,2 is ambiguous (G3): write cols=[1, 2]",
                      errors(ops(dedupe.format("cols=1,2"), version="VCL 0.4")))
        self.assertIn("unbalanced brackets: '[' without ']'", errors(ops(dedupe.format("cols=[1, 2"), version="VCL 0.4")))
        self.assertIn("a list [...] is only allowed as DEDUPE cols=", errors(ops("LET(n, [1, 2])", version="VCL 0.4")))
        for form in ("cols=1", "cols=[1, n + 1]"):
            with self.subTest(form):
                self.assertEqual(errors(ops(dedupe.format(form), version="VCL 0.4")), [])

    def test_key_columns_are_numbers(self):
        found = check_text(ops("DEDUPE(USED_RANGE($ws), cols=[1, CELL($ws, 1, 2)], header=yes)", version="VCL 0.4"))
        self.assertTrue(any("raw cell value" in m for _, m in found), found)


class Find(unittest.TestCase):
    def found(self, find):
        """The cell FIND marks with SET_VALUE(..., "hit"), or None when it gave NOTHING."""
        r = run(ops(f'IF NOT IS_NOTHING({find}) THEN SET_VALUE({find}, "hit")'))
        self.assertEqual(r.errors, 0)
        return next((rc for rc, v in cells(r).items() if v == "hit"), None)

    def test_row_by_row_and_the_first_match(self):
        self.assertEqual(self.found('FIND(USED_RANGE($ws), "b")'), (4, 1))         # rows 4 and 5 hold "b"
        self.assertEqual(self.found('FIND(USED_RANGE($ws), "KEY", match=part)'), (1, 1))
        self.assertEqual(self.found('FIND(RANGE($ws, "A2:B5"), "")'), None)       # no empty cell in A2:B5

    def test_row_order_not_column_order(self):
        grid = 'FIXTURE  t_01.xlsx sheet=Data cols=A:Key,B:Qty rows=2 data=A:values("x"x1,"y"x1)+boundary("x","y");B:values("y"x1,"x"x1)+boundary("y","x")'
        r = run(ops('SET_VALUE(FIND(RANGE($ws, "A2:B3"), "y"), "hit")').replace(FIXTURE, grid))
        self.assertEqual(cells(r)[(2, 2)], "hit")   # B2 comes before A3; column order would give A3

    def test_nothing(self):
        self.assertEqual(self.found('FIND(USED_RANGE($ws), "zz")'), None)
        self.assertEqual(self.found('FIND(USED_RANGE($ws), " ", match=part)'), None)   # nothing to look for
        self.assertEqual(run(ops("IF IS_NOTHING(NOTHING) THEN LET(n, 7)")).result, 7)
        r = run(ops('SET_VALUE(FIND(USED_RANGE($ws), "zz"), 1)'))                   # VBA error 91
        self.assertEqual((r.result, r.errors), (-2, 1))

    def test_a_range_variable_holds_it(self):
        text = ops('LET(hit, FIND(USED_RANGE($ws), "a"))', "IF NOT IS_NOTHING(hit) THEN LET(n, 1)",
                   head=["VAR      hit As Range"])
        self.assertEqual(run(text).result, 1)
        self.assertIn('Set hit = ForgeFind(ForgeUsedRange(ws), "a", True)', compile_text(text))

    def test_modifies_knows_the_sheet_of_find_in_place_only(self):
        self.assertEqual(errors(ops('SET_VALUE(FIND(USED_RANGE($ws), "a"), 1)')), [])
        found = errors(ops('LET(hit, FIND(USED_RANGE($ws), "a"))', "SET_VALUE(hit, 1)", head=["VAR      hit As Range"]))
        self.assertIn("SET_VALUE: MODIFIES can't tell which sheet this changes; write through CELL, ROW_RANGE, "
                      "USED_RANGE, RANGE or FIND of a $sheet argument", found)


def book_ops(*lines, head=(), modifies="MODIFIES $book sheets", version="VCL 0.4", call="P(t_01.xlsx, Data)"):
    """A small spec with a Workbook argument, around OP lines."""
    return "\n".join([version, "TASK", "PROC     P", "KIND     Function", "MODULE   modP",
                      "ARGS     book As Workbook, ws As Worksheet", "RET      Long", "VAR      n As Long", *head]
                     + ([modifies] if modifies else [])
                     + ["OP       " + o for o in lines]
                     + ["OP       RETURN(n)", "ERR      standard", FIXTURE,
                        f"TEST     fixture=t_01.xlsx call={call} assert=ERROR_COUNT = 0"])


class Sheets(unittest.TestCase):
    def test_add_sheet_and_its_place(self):
        r = run(book_ops('ADD_SHEET($book, name="Totals", after=$ws, into=out)', 'ADD_SHEET($book, name="First", into=s2)',
                         'WRITE_ROW(out, 1, "x")', 'IF SHEET_EXISTS($book, "totals") THEN LET(n, 1)',
                         head=["VAR      out As Worksheet", "VAR      s2 As Worksheet"]))
        self.assertEqual((r.result, r.errors, r.failed), (1, 0, []))
        self.assertEqual(list(r.state), ["Data", "Totals", "First"])   # after=$ws, then after the last sheet
        self.assertEqual(r.state["Totals"][0], [((1, 1), "x")])

    def test_never_a_replacement_and_never_a_bad_name(self):
        for name in ("DATA", "Q1/Q2", "History", " Pad"):   # a used name (case ignored), then three bad ones (S23)
            with self.subTest(name):
                r = run(book_ops(f'ADD_SHEET($book, name="{name}")'))   # run_tests doesn't validate
                self.assertEqual((r.result, r.errors), (-2, 1))
                self.assertEqual(list(r.state), ["Data"])                  # nothing was added

    def test_the_validator(self):
        def problems(text):   # leaves out the spec-wide TEST rules, which book_ops() breaks on purpose
            return [p.message for p in validator.validate_text(text) if p.severity == "error"
                    and "(rule 5)" not in p.message and "(rule 16)" not in p.message]
        self.assertEqual(problems(book_ops('ADD_SHEET($book, name="Totals")')), [])
        self.assertIn("ADD_SHEET adds a sheet to $book: declare MODIFIES $book sheets",
                      problems(book_ops('ADD_SHEET($book, name="Totals")', modifies=None)))
        self.assertIn("MODIFIES ... sheets is VCL 0.4: start the spec with 'VCL 0.4' to use it",
                      problems(book_ops('ADD_SHEET($book, name="Totals")', version="VCL 0.3")))
        self.assertIn("ADD_SHEET needs a Workbook argument ($book) as its first argument, so MODIFIES can name the "
                      "workbook it changes", problems(book_ops('ADD_SHEET(THIS_BOOK, name="Totals")')))
        self.assertIn("ADD_SHEET needs name=: Excel's own names (Sheet4) depend on the PC's language",
                      problems(book_ops("ADD_SHEET($book)")))
        self.assertTrue(any("is not a valid sheet name" in m for m in problems(book_ops('ADD_SHEET($book, name="Q1/Q2")'))))
        self.assertIn("TEST passes t_02.xlsx for $book, a Workbook: pass this TEST's fixture file, t_01.xlsx",
                      problems(book_ops('ADD_SHEET($book, name="Totals")', call="P(t_02.xlsx, Data)")))
        self.assertIn("MODIFIES $book: a Workbook argument gets new sheets only, so write MODIFIES $book sheets",
                      problems(book_ops('ADD_SHEET($book, name="Totals")', modifies="MODIFIES $book")))
        self.assertIn("SHEET_EXISTS is VCL 0.4: start the spec with 'VCL 0.4' to use it",
                      errors(ops('IF SHEET_EXISTS(THIS_BOOK, "Data") THEN LET(n, 1)')))
        found = problems(book_ops('ADD_SHEET($book, name="Totals", into=out)', "LET(out, $ws)", "SET_VALUE(CELL(out, 1, 1), 1)",
                                  head=["VAR      out As Worksheet"]))
        self.assertTrue(any(m.startswith("out holds the sheet made on line") for m in found), found)

    def test_copy_sheet_copies_values_fills_and_formats(self):
        r = run(book_ops("FORMAT(CELL($ws, 2, 1), fill=red)", 'COPY_SHEET($ws, name="Copy", into=c)',
                         "SET_VALUE(CELL(c, 1, 3), 1)", head=["VAR      c As Worksheet"],
                         modifies="MODIFIES $book sheets\nMODIFIES $ws fills"))
        self.assertEqual((r.errors, r.failed), (0, []))
        self.assertEqual(list(r.state), ["Data", "Copy"])
        data, copy = r.state["Data"], r.state["Copy"]
        self.assertEqual(copy[0], sorted(data[0] + [((1, 3), "1")]))   # the values, and the new one
        self.assertEqual(copy[1], data[1])                              # the fill
        for name in ("Data", "Copy?"):
            with self.subTest(name):
                r = run(book_ops(f'COPY_SHEET($ws, name="{name}")'))
                self.assertEqual((r.result, r.errors, list(r.state)), (-2, 1, ["Data"]))

    def test_copy_sheet_rules(self):
        def problems(text):
            return [p.message for p in validator.validate_text(text) if p.severity == "error"
                    and "(rule 5)" not in p.message and "(rule 16)" not in p.message]
        self.assertEqual(problems(book_ops('COPY_SHEET($ws, name="Copy")')), [])
        self.assertIn("COPY_SHEET adds a sheet to the workbook of the sheet it copies: give the spec a Workbook "
                      "argument $book and declare MODIFIES $book sheets",
                      problems(book_ops('COPY_SHEET($ws, name="Copy")', modifies="MODIFIES $ws")))
        self.assertIn("COPY_SHEET needs name=: Excel's own name for a copy (Template (2)) is not one a TEST can rely on",
                      problems(book_ops("COPY_SHEET($ws)")))
        found = check_text(book_ops('COPY_SHEET($book, name="Copy")'))
        self.assertTrue(any("COPY_SHEET: $book is not a sheet" in m for _, m in found), found)

    def test_an_added_sheet_needs_modifies_sheets_at_run_time_too(self):
        r = run(book_ops('ADD_SHEET($book, name="Totals")', modifies="MODIFIES $ws"))
        self.assertIn("UNCHANGED: sheet Totals was added", r.failed)


class Filter(unittest.TestCase):
    GRID = ('FIXTURE  t_01.xlsx sheet=Data cols=A:Key,B:Qty rows=4 data=A:values("a"x2,"b"x2)+boundary("a","b","a","b");'
            'B:values("1"x4)+boundary("1","1","1","1")')

    def run_grid(self, *lines, modifies="MODIFIES $ws"):
        text = ops(*lines, version="VCL 0.4").replace(FIXTURE, self.GRID)
        return run(text if modifies == "MODIFIES $ws" else text.replace("MODIFIES $ws", modifies))

    def test_hides_rows_both_ways_and_keeps_the_header(self):
        r = self.run_grid('FILTER(USED_RANGE($ws), col=1, crit="a")')
        self.assertEqual((r.errors, r.hidden), (0, {"Data": [3, 5]}))
        r = self.run_grid('FILTER(USED_RANGE($ws), col=1, crit="a")', 'FILTER(USED_RANGE($ws), col=1, crit=" B ")')
        self.assertEqual(r.hidden, {"Data": [2, 4]})                 # the second FILTER shows rows again
        r = self.run_grid('FILTER(USED_RANGE($ws), col=1, crit="zz")')
        self.assertEqual(r.hidden, {"Data": [2, 3, 4, 5]})           # row 1, the header, stays
        self.assertEqual(len(cells(r)), 10)                          # nothing deleted

    def test_row_ops_carry_hidden_rows(self):
        r = self.run_grid('FILTER(USED_RANGE($ws), col=1, crit="a")', "DELETE_ROW($ws, 2)", "INSERT_ROW($ws, 2)")
        self.assertEqual(r.hidden, {"Data": [3, 5]})                 # 3 -> 2 -> 3 and 5 -> 4 -> 5; row 2 is new
        r = self.run_grid('FILTER(USED_RANGE($ws), col=1, crit="a")', "INSERT_ROW($ws, 3)")
        self.assertEqual(r.hidden, {"Data": [4, 6]})                 # the new row 3 is visible

    def test_unchanged_sees_hidden_rows(self):
        r = self.run_grid('FILTER(USED_RANGE($ws), col=1, crit="a")', modifies="MODIFIES $ws fills")
        self.assertIn("UNCHANGED: Data row 3 was visible, now hidden", r.failed)

    def test_visible_count_and_the_rules(self):
        self.assertIn("assertion: VISIBLE_COUNT is VCL 0.4: start the spec with 'VCL 0.4' to use it",
                      [p.message for p in validator.validate_text(FILTER.replace("VCL 0.4", "VCL 0.3", 1))])
        both = FILTER.replace("OP       RETURN(n)", "OP       SORT(USED_RANGE($ws), key=1)\nOP       RETURN(n)")
        self.assertIn("a spec may not both FILTER and SORT: how Excel sorts rows that FILTER hid can't be proven "
                      "by a TEST (D42); sort in one spec and filter in another",
                      [p.message for p in validator.validate_text(both)])


class Unchanged(unittest.TestCase):
    def test_a_spec_that_writes_nothing_changes_nothing_from_0_2_on(self):   # D20, for every later version too
        for version, want in (("VCL 0.1", False), ("VCL 0.2", True), ("VCL 0.3", True), ("VCL 0.4", True)):
            with self.subTest(version):
                spec, problems = parse(ops("LET(n, 1)", version=version).replace("MODIFIES $ws\n", ""))
                self.assertEqual(problems, [])
                self.assertEqual(needs_unchanged_check(spec, modifies_of(spec)), want)


class Facts(unittest.TestCase):
    """An op that moves rows or writes a block forgets what the guards proved about that sheet."""
    GUARD = "IF IS_NUMBER(CELL($ws, 2, 2))"

    def test_the_type_checker_forgets(self):
        for op in ("DEDUPE(USED_RANGE($ws), cols=1, header=yes)", "WRITE_ARRAY(a, to=CELL($ws, 1, 1))"):
            with self.subTest(op):
                text = ops(self.GUARD, "  READ_ARRAY(CELL($ws, 1, 1), into=a)", "  " + op,
                           "  LET(n, CELL($ws, 2, 2) + 1)", head=["VAR      a As Variant"])
                found = check_text(text)
                self.assertTrue(any("raw cell value CELL($ws,2,2)" in m for _, m in found), found)

    def test_a_text_test_after_it_compiles(self):   # before the fix: "can't compare Num with Text"
        for op in ("DEDUPE(USED_RANGE($ws), cols=1, header=yes)", "WRITE_ARRAY(a, to=CELL($ws, 1, 1))"):
            with self.subTest(op):
                text = ops(self.GUARD, "  READ_ARRAY(CELL($ws, 1, 1), into=a)", "  " + op,
                           '  IF CELL($ws, 2, 2) = "x" THEN LET(n, 1)', head=["VAR      a As Variant"])
                self.assertIn('ForgeEq(ws.Cells(2, 2).Value, "x")', compile_text(text))


class ForgeLib(unittest.TestCase):
    def body(self, name):
        for kind in ("Function", "Sub"):
            head = f"Public {kind} {name}("
            if head in FORGELIB:
                start = FORGELIB.index(head)
                return FORGELIB[start:FORGELIB.index("\nEnd ", start)]
        raise AssertionError(f"{name} is not in ForgeLib")

    def test_the_new_helpers(self):
        self.assertIn("ws.Rows(rowNo).ClearFormats", self.body("ForgeInsertRow"))
        self.assertIn("If rng.Cells.CountLarge = 1 Then", self.body("ForgeReadArray"))
        self.assertIn('ForgeFileName = Mid$(s, InStrRev(s, "\\") + 1)', self.body("ForgeFileName"))
        dedupe = self.body("ForgeDedupe")
        self.assertLess(dedupe.index("key = ForgeRowKey(parts)"), dedupe.index("Step -1"))   # marked first, deleted after
        self.assertIn("If doomed(r) Then rng.Worksheet.Rows(rng.Row + r - 1).Delete", dedupe)
        self.assertIn("If ForgeIsEmpty(values(i)) Or ForgeIsError(values(i)) Then", self.body("ForgeRowKey"))
        add = self.body("ForgeAddSheet")
        self.assertLess(add.index("already exists"), add.index("book.Worksheets.Add(After:=after)"))   # checked first
        self.assertIn("If Not ForgeSheetNameOk(s) Then Err.Raise 1004", add)
        flt = self.body("ForgeFilter")
        self.assertNotIn("AutoFilter", flt.split("\n", 4)[-1])     # below the comment: no AutoFilter
        self.assertIn("ws.Rows(r).Hidden = Not ForgeSame(ws.Cells(r, col).Value, crit)", flt)
        last = self.body("ForgeLastRow")
        self.assertIn("To last + 1 Step -1", last)                  # rows below End(xlUp) are checked too
        self.assertIn("ws.Rows(rowNo).Hidden = False", self.body("ForgeInsertRow"))
        copy = self.body("ForgeCopySheet")
        self.assertLess(copy.index("already exists"), copy.index("src.Copy After:=after"))   # checked first
        find = self.body("ForgeFind")
        self.assertIn("For Each cell In rng.Cells", find)
        self.assertNotIn(".Find(", find)   # Range.Find keeps the settings of the last search
        self.assertIn("If Not whole And Len(want) = 0 Then Exit Function", find)


if __name__ == "__main__":
    unittest.main()
