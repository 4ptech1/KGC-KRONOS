"""VEP 0.4 (docs/vep-0.4-rev3-rest.md). Phase 1: the interpreter runs what the compiler already
compiled: INSERT_ROW, READ_ARRAY, WRITE_ARRAY, ARR, ROWS, COLS, SELECT_CASE, LOG, MSG, RANGE, SHEET,
FILE_NAME and FOR_EACH over SHEETS and CELLS (D36, D37). Phase 2: the VCL 0.4 line, lists in
DEDUPE cols=[...] and DEDUPE itself (S21, D38). Phase 3: FIND, IS_NOTHING and NOTHING (S22, D39). Phase 4a: Workbook
arguments, MODIFIES $book sheets, ADD_SHEET and SHEET_EXISTS (S23, D40). Phase 4b: COPY_SHEET (D41). Phase 5:
FILTER, hidden rows and VISIBLE_COUNT (D42). Phase 6: SET_FORMULA, the formula check and FORMULA (D43).
Phase 7a: LIST_FILES and ITEMS (S24, D44). Phase 7b: READ_TEXT, WRITE_TEXT, file texts and FILE_TEXT
(S25, D45). The switch of spec.md and the prompt to VCL 0.4 (E12, D46)."""
import datetime
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "compiler"))
import validator  # noqa: E402
sys.path.insert(0, str(ROOT / "vcl"))
import semantics as S  # noqa: E402
from vba_lint import lint  # noqa: E402
from vcl_codegen import CompileError, assemble, compile_text  # noqa: E402
from vcl_formula import formula_problems, sheet_names  # noqa: E402
from vcl_fixtures import SANDBOX, MockFS  # noqa: E402
from vcl_interp import VclError, modifies_of, needs_unchanged_check, read_text, run_tests, write_text  # noqa: E402
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

# SET_FORMULA (D43): the same R1C1 formula in every Total cell, checked by its text.
FORMULA = r"""VCL 0.4
TASK
# ASSUME: each Total cell gets the formula Qty * Rate, written in R1C1 style so every row has the same text
# ASSUME: the rows are 2 to the last Item row; the result is how many formulas were written
PROC     AddLineTotals
KIND     Function
MODULE   modTotals
ARGS     ws As Worksheet
RET      Long
VAR      cTotal As Long
VAR      n As Long
PRE      COL($ws, "Item") > 0; COL($ws, "Qty") > 0; COL($ws, "Rate") > 0; COL($ws, "Total") > 0
MODIFIES $ws["Total"]
OP       LET(cTotal, COL($ws, "Total"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, COL($ws, "Item")) var=r
OP         SET_FORMULA(CELL($ws, r, cTotal), "=RC[-2]*RC[-1]", style=r1c1)
OP         LET(n, n + 1)
OP       RETURN(n)
ERR      standard
FIXTURE  lines_01.xlsx sheet=Lines cols=A:Item,B:Qty,C:Rate,D:Total rows=5 data=A:seq(1,1);B:int(1,9);C:int(10,99)
FIXTURE  lines_01.xlsx sheet=Bare cols=A:Item,B:Qty,C:Rate,D:Total rows=0
FIXTURE  lines_01.xlsx sheet=Other cols=A:Item rows=1 data=A:seq(1,1)
TEST     fixture=lines_01.xlsx call=AddLineTotals(Lines) assert=RESULT = 5 AND ERROR_COUNT = 0 AND FORMULA(Lines!D2) = "=RC[-2]*RC[-1]" AND FORMULA(Lines!D6) = "=RC[-2]*RC[-1]" AND COUNTA(Lines!D:D) = 6 AND FORMULA(Lines!D7) = ""
TEST     fixture=lines_01.xlsx call=AddLineTotals(Lines) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTA(Lines!B:B) = PRE_COUNTA(Lines!B:B)
TEST     fixture=lines_01.xlsx call=AddLineTotals(Bare) assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTA(Bare!D:D) = 1
TEST     fixture=lines_01.xlsx call=AddLineTotals(Other) assert=RESULT = -1 AND ERROR_COUNT = 1
"""

# LIST_FILES (D44): the PDFs directly in the dump folder, in S24 order.
LIST_FILES = r"""VCL 0.4
TASK
# ASSUME: the report lists the PDF files directly in the dump folder, not in its sub-folders, in name order
# ASSUME: the names go under the File header from row 2; the result is how many there are
PROC     ListDumpPdfs
KIND     Function
MODULE   modDump
ARGS     ws As Worksheet, dump As String
RET      Long
VAR      files As Collection
VAR      r As Long
PRE      FOLDER_EXISTS($dump); COL($ws, "File") > 0
MODIFIES $ws["File"]
OP       LIST_FILES($dump, "*.pdf", into=files)
OP       LET(r, 1)
OP       FOR_EACH var=fileName in=ITEMS(files)
OP         LET(r, r + 1)
OP         SET_VALUE(CELL($ws, r, COL($ws, "File")), fileName)
OP       RETURN(COUNT(files))
ERR      standard
FIXTURE  dump_01.xlsx sheet=Report cols=A:File rows=0
FIXTURE  dump_01.xlsx sheet=Other cols=A:Item rows=1 data=A:seq(1,1)
FIXTURE  dump_01.xlsx folder=Dump files=b.pdf,A.PDF,c.txt,x.pdfa,sub\d.pdf
FIXTURE  dump_01.xlsx folder=Empty
TEST     fixture=dump_01.xlsx call=ListDumpPdfs(Report, Dump) assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(Report!A2) = "A.PDF" AND VALUE(Report!A3) = "b.pdf" AND COUNTA(Report!A:A) = 3 AND FILE_COUNT("Dump") = 5
TEST     fixture=dump_01.xlsx call=ListDumpPdfs(Report, Dump) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND COUNTA(Report!A:A) = PRE_COUNTA(Report!A:A) + 2
TEST     fixture=dump_01.xlsx call=ListDumpPdfs(Report, Empty) assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTA(Report!A:A) = 1
TEST     fixture=dump_01.xlsx call=ListDumpPdfs(Other, Dump) assert=RESULT = -1 AND ERROR_COUNT = 1
"""

# READ_TEXT and WRITE_TEXT (D45): every status a file can give, and new files only.
TEXT_FILES = r"""VCL 0.4
TASK
# ASSUME: each job folder may hold job.txt with the job's title, which goes to the Title column (or the status)
# ASSUME: done.txt ("done: <title>") is written next to it when the title was read; its status goes to Note
PROC     MarkJobsDone
KIND     Function
MODULE   modNotes
ARGS     ws As Worksheet, jobs As String
RET      Long
VAR      job As String
VAR      title As String
VAR      readStatus As String
VAR      writeStatus As String
VAR      n As Long
PRE      FOLDER_EXISTS($jobs); COL($ws, "Job") > 0; COL($ws, "Title") > 0; COL($ws, "Note") > 0
MODIFIES $ws["Title", "Note"]
MODIFIES $jobs files
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, COL($ws, "Job")) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, COL($ws, "Job")))
OP         LET(job, TRIM(CELL($ws, r, COL($ws, "Job"))))
OP         READ_TEXT(PATH_JOIN($jobs, SAFE_NAME(job), "job.txt"), title, into=readStatus)
OP         SET_VALUE(CELL($ws, r, COL($ws, "Title")), IIF(readStatus = "Read", title, readStatus))
OP         IF readStatus = "Read"
OP           WRITE_TEXT(PATH_JOIN($jobs, SAFE_NAME(job), "done.txt"), "done: " & title, into=writeStatus)
OP           SET_VALUE(CELL($ws, r, COL($ws, "Note")), writeStatus)
OP           IF writeStatus = "Written" THEN LET(n, n + 1)
OP       RETURN(n)
ERR      standard
FIXTURE  jobs_01.xlsx sheet=Register cols=A:Job,B:Title,C:Note rows=6 data=A:values("J1"x1,"J2"x1,"J3"x1,"J4"x1,"J5"x1,""x1)+boundary("J1","","J2","J3","J4","J5")
FIXTURE  jobs_01.xlsx sheet=Other cols=A:Item rows=1 data=A:seq(1,1)
FIXTURE  jobs_01.xlsx folder=Jobs files=J1\job.txt:"Shell 1200",J2\job.txt:"Head",J2\done.txt:"done: Head",J3\other.txt:"x",J4\job.txt:"Nozzle",J5\job.txt:"Lock" locked=J5\job.txt
TEST     fixture=jobs_01.xlsx call=MarkJobsDone(Register, Jobs) assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(Register!B2) = "Shell 1200" AND VALUE(Register!B5) = "Missing" AND VALUE(Register!B7) = "Locked" AND VALUE(Register!C4) = "Already there" AND VALUE(Register!C6) = "Written" AND FILE_TEXT("Jobs\J1\done.txt") = "done: Shell 1200" AND FILE_EXISTS("Jobs\J4\done.txt")
TEST     fixture=jobs_01.xlsx call=MarkJobsDone(Register, Jobs) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND FILE_COUNT("Jobs") = 8 AND FILE_TEXT("Jobs\J4\done.txt") = "done: Nozzle" AND COUNTA(Register!A:A) = PRE_COUNTA(Register!A:A)
TEST     fixture=jobs_01.xlsx call=MarkJobsDone(Other, Jobs) assert=RESULT = -1 AND ERROR_COUNT = 1
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
                           ("ADD_SHEET", ADD_SHEET), ("COPY_SHEET", COPY_SHEET), ("FILTER", FILTER),
                           ("FORMULA", FORMULA), ("LIST_FILES", LIST_FILES), ("TEXT_FILES", TEXT_FILES)):
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
        vba = compile_text(FORMULA)
        self.assertEqual(lint(assemble([vba])), [])
        self.assertIn('ForgeSetFormula ws.Cells(r, cTotal), "=RC[-2]*RC[-1]", True', vba)
        vba = compile_text(LIST_FILES)
        self.assertEqual(lint(assemble([vba])), [])
        self.assertIn('Set files = ForgeListFiles(dump, dump, "*.pdf")', vba)
        self.assertIn("For Each fileName In files", vba)
        vba = compile_text(TEXT_FILES)
        self.assertEqual(lint(assemble([vba])), [])
        self.assertIn('readStatus = ForgeReadText(jobs, ForgePathJoin(jobs, Array(ForgeSafeName(job), "job.txt")), title)', vba)
        self.assertIn('writeStatus = ForgeWriteText(jobs, ForgePathJoin(jobs, Array(ForgeSafeName(job), "done.txt")), '
                      '("done: " & title))', vba)

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
            and "(rule 5)" not in p.message and "(rule 16)" not in p.message]


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


class Formulas(unittest.TestCase):
    def test_the_excel_2016_list(self):
        for ok in ("=SUM(B2:B9)", '=IFERROR(VLOOKUP(A2,Rates!A:B,2,FALSE),"none")', "='Rev (2)'!A1+1",
                   "=IF(A1=#N/A,1,0)", "=sumifs(c:c,a:a,\"x\")"):
            with self.subTest(ok):
                self.assertEqual(formula_problems(ok), [])
        self.assertEqual(formula_problems("=RC[-2]*RC[-1]", r1c1=True), [])
        for bad, why in (("=XLOOKUP(A2,B:B,C:C)", "XLOOKUP is not on"), ("=TEXTJOIN(\",\",TRUE,A1:A3)", "TEXTJOIN is not on"),
                         ("=LET(x,1,x)", "LET is not on"), ('=INDIRECT("A1")', "INDIRECT is not on"),
                         ('=TEXT(A1,"yyyy")', "TEXT is not on"), ("=[Book1.xlsx]Sheet1!A1", "another workbook"),
                         ("=Table1[Qty]", "or a table"), ("=RC[-2]", "another workbook"), ("=@A1:A5", "@ is"),
                         ("=A1#", "spill"), ("=_xlfn.CONCAT(A1,B1)", "_xlfn."), ("SUM(A1)", "starts with =")):
            with self.subTest(bad):
                self.assertTrue(any(why in p for p in formula_problems(bad)), formula_problems(bad))

    def test_formula_cells_have_no_value_in_the_interpreter(self):
        set_c2 = 'SET_FORMULA(CELL($ws, 2, 3), "=B2*2")'
        self.assertEqual(run(ops(set_c2, "LET(n, 1)")).errors, 0)
        r = run(ops(set_c2, "IF IS_EMPTY(CELL($ws, 2, 3)) THEN LET(n, 1)"))
        self.assertEqual((r.result, r.errors), (-2, 1))                          # read as a value: stops
        r = run(ops('SET_FORMULA(ROW_RANGE($ws, 2, 3, 4), "=B2*2")'))          # a1 on two cells
        self.assertEqual((r.result, r.errors), (-2, 1))
        r = run(ops('SET_FORMULA(ROW_RANGE($ws, 2, 3, 4), "=RC[-1]*2", style=r1c1)'))
        self.assertEqual((cells(r)[(2, 3)], cells(r)[(2, 4)]), ("=RC[-1]*2", "=RC[-1]*2"))

    def test_the_validator(self):
        found = errors(ops('SET_FORMULA(CELL($ws, 2, 3), "=XLOOKUP(A2,B:B,C:C)")'))
        self.assertIn("SET_FORMULA: XLOOKUP is not on the Excel 2016 list of SET_FORMULA functions (D43)", found)
        found = errors(ops('SET_FORMULA(CELL($ws, 2, 3), "=B2*2")', "DELETE_ROW($ws, 3)"))
        self.assertIn("a spec may not both SET_FORMULA and DELETE_ROW: Excel rewrites formula references when rows "
                      "move, which a TEST can't prove (D43)", found)
        self.assertIn("assertion: FORMULA is VCL 0.4: start the spec with 'VCL 0.4' to use it",
                      [p.message for p in validator.validate_text(FORMULA.replace("VCL 0.4", "VCL 0.3", 1))])


class ListFiles(unittest.TestCase):
    def listed(self, folder_expr, pattern="*.pdf", pre="FOLDER_EXISTS($dump); COL($ws, \"File\") > 0"):
        text = LIST_FILES.replace("LIST_FILES($dump, \"*.pdf\"", f'LIST_FILES({folder_expr}, "{pattern}"')
        text = text.replace("FIXTURE  dump_01.xlsx folder=Empty",
                            "FIXTURE  dump_01.xlsx folder=Empty files=Jobs\\J1\\r1.pdf,Jobs\\J1\\r2.PDF,Jobs\\J1\\n.txt")
        r = run_tests(text)[2]   # the TEST on the Empty folder, which now has files below it
        return r.result

    def test_a_folder_below_the_root(self):
        self.assertEqual(self.listed('PATH_JOIN($dump, "Jobs", "J1")'), 2)
        self.assertEqual(self.listed('PATH_JOIN($dump, "Jobs")'), 0)          # files below, not in it
        self.assertEqual(self.listed('PATH_JOIN($dump, "Nope")'), 0)          # a missing folder: an empty list
        self.assertEqual(self.listed('PATH_JOIN($dump, "Jobs", "J1")', pattern="*"), 3)

    def test_the_rules(self):
        found = [p.message for p in validator.validate_text(LIST_FILES.replace('"*.pdf"', '"..\\*.pdf"'))]
        self.assertTrue(any("LIST_FILES pattern" in m for m in found), found)
        found = [p.message for p in validator.validate_text(LIST_FILES.replace("VCL 0.4", "VCL 0.3", 1))]
        self.assertIn("LIST_FILES is VCL 0.4: start the spec with 'VCL 0.4' to use it", found)
        found = [p.message for p in validator.validate_text(LIST_FILES.replace("PRE      FOLDER_EXISTS($dump); ", "PRE      "))]
        self.assertTrue(any("FOLDER_EXISTS($dump)" in m and "rule 14" in m for m in found), found)
        found = check_text(LIST_FILES.replace("VAR      files As Collection", "VAR      files As String"))
        self.assertTrue(any("but the names are a list: declare it As Collection" in m for _, m in found), found)


class TextFiles(unittest.TestCase):
    ROOT = SANDBOX + "\\Jobs"

    def fs(self):
        fs = MockFS()
        fs.add_folder(self.ROOT)
        fs.add_folder(self.ROOT + "\\J1")
        fs.add_file(self.ROOT + "\\J1\\a.txt", 6, text="Head\r\n")
        fs.add_file(self.ROOT + "\\J1\\cafe.txt", 4, text="caf\u00e9")
        fs.add_file(self.ROOT + "\\J1\\big.txt", 32768)
        fs.add_file(self.ROOT + "\\J1\\held.txt", 4, locked=True, text="Lock")
        return fs

    def test_read_text_in_status_order(self):
        fs, root = self.fs(), self.ROOT
        for path, want in (("", ("Bad name", "")), (root + "\\J1\\a*.txt", ("Bad name", "")),
                           (root + "\\" + "x" * 260, ("Too long", "")), (SANDBOX + "\\Other\\a.txt", ("Bad name", "")),
                           (root + "\\J1\\none.txt", ("Missing", "")), (root + "\\J1", ("Missing", "")),
                           (root + "\\J1\\big.txt", ("Too big", "")), (root + "\\J1\\held.txt", ("Locked", "")),
                           (root + "\\J1\\cafe.txt", ("Not text", "")), (root + "\\J1\\a.txt", ("Read", "Head"))):
            with self.subTest(path[-20:]):
                self.assertEqual(read_text(fs, root, path), want)

    def test_write_text_new_files_only(self):
        fs, root = self.fs(), self.ROOT
        for path, text, want in (("", "x", "Bad name"), (root + "\\" + "x" * 260, "x", "Too long"),
                                 (SANDBOX + "\\Other\\n.txt", "x", "Bad name"), (root + "\\J1\\n.txt", "\u20ac", "Not text"),
                                 (root + "\\J9\\n.txt", "x", "No folder"), (root + "\\J1", "x", "Conflict"),
                                 (root + "\\J1\\a.txt", "Head", "Conflict"),          # 4 bytes, not 6
                                 (root + "\\J1\\a.txt", "Head!\n", "Conflict"),      # 6 bytes, other text
                                 (root + "\\J1\\held.txt", "Lock", "Locked"), (root + "\\J1\\a.txt", "Head\r\n", "Already there"),
                                 (root + "\\J1\\n.txt", "new", "Written"), (root + "\\J1\\n.txt", "new", "Already there")):
            with self.subTest(want):
                self.assertEqual(write_text(fs, root, path, text), want)
        self.assertEqual(fs.text_of(root + "\\J1\\a.txt"), "Head\r\n")       # never written over

    def test_s25_length(self):
        self.assertTrue(S.text_ok("x" * 32767))
        self.assertFalse(S.text_ok("x" * 32768))
        self.assertEqual(S.file_text("x" * 32767 + "\n"), None)    # the raw file is one too many
        self.assertIn("Private Const FORGE_MAX_TEXT As Long = 32767", FORGELIB)

    def test_the_rules(self):
        def problems(text):
            return [p.message for p in validator.validate_text(text) if p.severity == "error"]
        self.assertIn("READ_TEXT needs a declared String VAR second, which gets the file's text, got '\"x\"'",
                      problems(TEXT_FILES.replace("job.txt\"), title, into", 'job.txt"), "x", into')))
        self.assertIn("WRITE_TEXT has no option mode=",
                      problems(TEXT_FILES.replace("title, into=writeStatus", "title, mode=new, into=writeStatus")))
        self.assertIn("WRITE_TEXT adds to $jobs, but no line declares MODIFIES $jobs files",
                      problems(TEXT_FILES.replace("MODIFIES $jobs files\n", "")))
        self.assertIn("READ_TEXT is VCL 0.4: start the spec with 'VCL 0.4' to use it",
                      problems(TEXT_FILES.replace("VCL 0.4", "VCL 0.3", 1)))
        self.assertTrue(any("must be one line of printable ASCII" in m
                            for m in problems(TEXT_FILES.replace('"Head"', '"H\u00e9ad"'))))
        unlocked = TEXT_FILES.replace(" locked=J5\\job.txt", "")   # rule 19 (D46)
        self.assertNotEqual(unlocked, TEXT_FILES)
        self.assertIn(("warning", "the spec reads files, but no FIXTURE folder has a locked= file, so no TEST shows "
                                  "what happens to a file that is open elsewhere (rule 19)"),
                      [(p.severity, p.message) for p in validator.validate_text(unlocked)])


class Unchanged(unittest.TestCase):
    def test_a_spec_that_writes_nothing_changes_nothing_from_0_2_on(self):   # D20, for every later version too
        for version, want in (("VCL 0.1", False), ("VCL 0.2", True), ("VCL 0.3", True), ("VCL 0.4", True)):
            with self.subTest(version):
                spec, problems = parse(ops("LET(n, 1)", version=version).replace("MODIFIES $ws\n", ""))
                self.assertEqual(problems, [])
                self.assertEqual(needs_unchanged_check(spec, modifies_of(spec)), want)


class VersionSwitch(unittest.TestCase):
    """E12, D46: every committed spec and the prompt's examples, read as VCL 0.4, give the same
    problems and the same TEST results as under their own line (the examples were 0.3)."""

    def test_every_spec_and_example_gives_the_same_results_under_0_4(self):
        paths = sorted((ROOT / "specs").glob("*.vcl")) + [ROOT / "vcl" / "examples" / name
                                                         for name in ("example1_rev3.vcl", "example3_files.vcl")]
        day = datetime.date(2026, 10, 20)
        for path in paths:
            first, rest = path.read_text(encoding="utf-8").split("\n", 1)
            with self.subTest(spec=path.name):
                self.assertIn(first, ("VCL 0.2", "VCL 0.3", "VCL 0.4"))
                older = ("VCL 0.3" if first == "VCL 0.4" else first) + "\n" + rest
                newer = "VCL 0.4\n" + rest
                self.assertEqual(*[[(p.line, p.severity, p.message) for p in validator.validate_text(t)]
                                   for t in (older, newer)])
                self.assertEqual(*[[(r.line, r.passed, r.result, r.errors, r.state, r.files, r.hidden, r.texts)
                                    for r in run_tests(t, day)] for t in (older, newer)])

    def test_real_workbooks_are_refused_until_march_2027(self):   # D36: the language keeps them
        for op in ('OPEN_BOOK("x.xlsx", readonly=yes, into=b)', 'SAVE_AS(THIS_BOOK, "out.xlsx", format=xlsx)',
                   "CLOSE_BOOK(THIS_BOOK, save=no)"):
            with self.subTest(op):
                text = ops(op, head=["VAR      b As Workbook"], version="VCL 0.4")
                why = f"{op.split('(')[0]} opens or saves a real workbook: refused until March 2027 (D36)"
                with self.assertRaises(CompileError) as caught:
                    compile_text(text)
                self.assertEqual(str(caught.exception), why)
                self.assertEqual(run(text).log[0][2], "runtime error: " + why)

    def test_a_build_says_why_it_cannot_be_tested(self):   # D51: it used to say "spec doesn't parse: []"
        # A BUILD before VCL 0.5 has no MODIFIES in its TASKs, so its TESTs can't run (D49, D52);
        # tests/test_vcl05.py runs a 0.5 BUILD's TESTs.
        sys.path.insert(0, str(ROOT / "tests"))
        from test_parser import MINIMAL_BUILD
        self.assertEqual([p for p in validator.validate_text(MINIMAL_BUILD) if p.severity == "error"], [])
        with self.assertRaises(VclError) as caught:
            run_tests(MINIMAL_BUILD)
        self.assertEqual(str(caught.exception),
                         "a BUILD runs its TESTs from VCL 0.5 on: before, its TASKs don't declare MODIFIES, so no "
                         "TEST could check that nothing else changed (D49)")


class ReviewPass(unittest.TestCase):
    """D51: what the second review pass after item 2 found in ForgeLib, the formula check and SORT."""

    def test_forge_sort_gives_every_sticky_option(self):
        # Range.Sort reuses the sheet's last Orientation, OrderCustom and MatchCase when they are left out.
        start = FORGELIB.index("Public Sub ForgeSort(")
        body = FORGELIB[start:FORGELIB.index("\nEnd Sub", start)]
        for option in ("Orientation:=xlTopToBottom", "OrderCustom:=1", "MatchCase:=False", "DataOption1:=xlSortNormal"):
            self.assertIn(option, body)

    def test_list_files_leaves_out_hidden_and_system_files(self):   # ~$ owner files, Thumbs.db
        start = FORGELIB.index("Public Function ForgeListFiles(")
        self.assertIn("If (f.Attributes And 6) = 0 Then", FORGELIB[start:FORGELIB.index("\nEnd Function", start)])

    def test_sort_puts_errors_after_booleans_and_dates_by_serial(self):
        from vcl_fixtures import Sheet
        from vcl_interp import Machine, Rng
        sh = Sheet("Data")
        column = ["Key", S.XlError("#N/A"), 5, datetime.datetime(2026, 1, 1, 15), "b",
                  datetime.datetime(2026, 1, 1, 9), True, None, 46023, S.XlError("#DIV/0!")]
        for r, v in enumerate(column, start=1):
            if v is not None:
                sh.set(r, 1, v)
            sh.set(r, 2, r)
        Machine.sort(None, Rng(sh, 1, 1, len(column), 2), 1, True, True)
        # numbers and dates by Excel's serial (46023 is 1 Jan 2026), text, booleans, errors (kept in
        # order), blanks; errors used to sort by their memory address, and times were ignored
        self.assertEqual([sh.get(r, 2) for r in range(2, len(column) + 1)], [3, 9, 6, 4, 5, 7, 2, 10, 8])

    def test_a_formula_names_no_file(self):
        for bad in ("='C:\\x\\[Other.xlsx]Data'!A1", "='\\\\srv\\share\\[a.xlsx]S'!B2", "='a/b'!A1"):
            with self.subTest(bad):
                self.assertTrue(any("names another file" in p for p in formula_problems(bad)))
        self.assertEqual(formula_problems("='Weld Log'!A1+'O''Neil'!B2"), [])
        self.assertEqual(sheet_names("=SUM(Data!B2:B9)+'O''Neil'!A1+Data!C1"), ["Data", "O'Neil"])
        self.assertEqual(sheet_names('="x!"&A1&RC[-1]'), [])

    def test_a_formula_names_only_sheets_the_test_workbook_has(self):
        text = ops('SET_FORMULA(CELL($ws, 2, 3), "=SUM(Data!A2:A3)+\'NDE Register\'!A1")', version="VCL 0.4")
        found = [p.message for p in validator.validate_text(text) if p.severity == "error"]
        self.assertIn("SET_FORMULA names the sheet NDE Register, which the TEST workbook t_01.xlsx doesn't have: "
                      "in Excel the formula would ask for a file (D51)", found)
        self.assertFalse(any("the sheet Data," in m for m in found))
        added = ops('ADD_SHEET(THIS_BOOK, name="NDE Register", into=s)',
                    'SET_FORMULA(CELL($ws, 2, 3), "=\'NDE Register\'!A1")', head=["VAR      s As Worksheet"],
                    version="VCL 0.4")
        self.assertFalse(any("SET_FORMULA names the sheet" in p.message for p in validator.validate_text(added)))


class VbaReviewFixes(unittest.TestCase):
    """D48: what the ForgeLib review found, fixed in ForgeLib and, where it changes a meaning, in the
    interpreter too."""

    def body(self, name):
        return forgelib_body(name)

    def test_set_formula_refuses_a_cell_formatted_as_text(self):
        r = run(ops('FORMAT(CELL($ws, 2, 3), numfmt="@")', 'SET_FORMULA(CELL($ws, 2, 3), "=1+1")', version="VCL 0.4"))
        self.assertEqual((r.result, r.errors), (-2, 1))
        self.assertIn("formatted as Text", r.log[0][2])
        ok = run(ops('SET_FORMULA(CELL($ws, 2, 3), "=1+1")', version="VCL 0.4"))
        self.assertEqual((ok.result, ok.errors), (0, 0))
        body = self.body("ForgeSetFormula")
        self.assertLess(body.index('c.NumberFormat = "@"'), body.index("rng.Formula"))   # checked before writing

    def test_text_file_paths_are_printable_ascii(self):
        root = SANDBOX + "\\Jobs"
        fs = MockFS()
        fs.add_folder(root)
        fs.add_file(root + "\\a.txt", 4, text="Head")
        self.assertEqual(read_text(fs, root, root + "\\a.txt"), ("Read", "Head"))
        self.assertEqual(read_text(fs, root, root + "\\\u03a9.txt"), ("Bad name", ""))
        self.assertEqual(write_text(fs, root, root + "\\\u00e9.txt", "x"), "Bad name")
        for name in ("ForgeReadText", "ForgeWriteText"):
            self.assertIn("Or Not ForgeIsAsciiPath(path) Then", self.body(name))

    def test_sheet_ops_never_reach_another_workbook(self):
        for name, change in (("ForgeAddSheet", "book.Worksheets.Add(After:=after)"), ("ForgeCopySheet", "src.Copy After:=after")):
            with self.subTest(name):
                body = self.body(name)
                self.assertLess(body.index("ElseIf Not after.Parent Is book Then"), body.index(change))

    def test_vba_keeps_the_meanings_of_the_python_reference(self):
        self.assertIn("(AscW(Mid$(s, i, 1)) And &HFFFF&) < 32", self.body("ForgeSheetNameOk"))
        self.assertIn("k = LCase$(ForgeText(values(i)))", self.body("ForgeRowKey"))
        self.assertIn("vbBinaryCompare", self.body("ForgeSameRowKey"))
        self.assertIn("seen.CompareMode = vbBinaryCompare", self.body("ForgeDedupe"))
        self.assertFalse(S.same_row_key(["Stra\u00dfe"], ["STRASSE"]))
        self.assertTrue(S.sheet_name_ok("\ubcf4\uace0\uc11c"))


class FailureMessages(unittest.TestCase):
    """A failing TEST names each false assertion and what each side measured (review pass, D47)."""

    def test_a_false_assertion_says_what_it_measured(self):
        text = (ROOT / "specs" / "task_01.vcl").read_text(encoding="utf-8")
        bad = text.replace("COUNTA(Shortages![Part No]) = RESULT + 1", "COUNTA(Shortages![Part No]) = RESULT - (1 + 1)", 1)
        self.assertNotEqual(bad, text)
        failed = [f for r in run_tests(bad, datetime.date(2026, 10, 20)) for f in r.failed]
        self.assertEqual(len(failed), 1)
        self.assertRegex(failed[0], r"^ASSERT: COUNTA\(Shortages!\[Part No\]\) = RESULT - \(1 \+ 1\) is false "
                                    r"\(left (\d+), right (\d+)\)$")
        left, right = map(int, re.search(r"left (\d+), right (\d+)", failed[0]).groups())
        self.assertEqual(left - right, 3)   # RESULT + 1 holds, so the sides differ by 3

    def test_an_assertion_that_cant_be_checked_says_why(self):
        r = run(ops("LET(n, 1)").replace("assert=ERROR_COUNT = 0", "assert=VALUE(Data![Nope]2) = 1"))
        self.assertEqual(r.failed, ["ASSERT: VALUE(Data![Nope]2) = 1 can't be checked: no header [Nope] on Data"])


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


def forgelib_body(name):
    """The text of one ForgeLib procedure, from its header to its End line."""
    for kind in ("Function", "Sub"):
        head = f"Public {kind} {name}("
        if head in FORGELIB:
            start = FORGELIB.index(head)
            return FORGELIB[start:FORGELIB.index("\nEnd ", start)]
    raise AssertionError(f"{name} is not in ForgeLib")


class ForgeLib(unittest.TestCase):
    def body(self, name):
        return forgelib_body(name)

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
        self.assertNotIn("FormulaLocal", self.body("ForgeSetFormula").split("\n", 3)[-1])
        flt = self.body("ForgeFilter")
        self.assertNotIn("AutoFilter", flt.split("\n", 4)[-1])     # below the comment: no AutoFilter
        self.assertIn("ws.Rows(r).Hidden = Not ForgeSame(ws.Cells(r, col).Value, crit)", flt)
        last = self.body("ForgeLastRow")
        self.assertIn("To last + 1 Step -1", last)                  # rows below End(xlUp) are checked too
        self.assertIn("ws.Rows(rowNo).Hidden = False", self.body("ForgeInsertRow"))
        listing = self.body("ForgeListFiles")
        self.assertIn("ForgeFso().GetFolder(folder).Files", listing)
        self.assertNotIn("Dir(", listing)
        self.assertNotIn(" Like ", self.body("ForgeNameMatches"))
        write = self.body("ForgeWriteText")
        self.assertIn("Set ts = fso.CreateTextFile(path, False, False)", write)   # never over a file; ANSI
        self.assertLess(write.index("ElseIf fso.FileExists(path) Then"), write.index("CreateTextFile"))
        self.assertIn("Close #f\n    Err.Raise errNo", FORGELIB[FORGELIB.index("Private Function ForgeReadRaw("):])
        copy = self.body("ForgeCopySheet")
        self.assertLess(copy.index("already exists"), copy.index("src.Copy After:=after"))   # checked first
        find = self.body("ForgeFind")
        self.assertIn("For Each cell In rng.Cells", find)
        self.assertNotIn(".Find(", find)   # Range.Find keeps the settings of the last search
        self.assertIn("If Not whole And Len(want) = 0 Then Exit Function", find)


if __name__ == "__main__":
    unittest.main()
