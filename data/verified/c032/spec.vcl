VCL 0.6
BUILD
# ASSUME: the real sheets are "Weld Log", "Weld Exceptions" and "Welder Exceptions" and all three exist; the TESTs call them Log, Exc and ByWelder
# ASSUME: Weld Log headers (row 1) include "Weld Seam No", "Welder ID" and "Visual"; the row-1 headings of Weld Exceptions are copied from the log
# ASSUME: a row with a Weld Seam No is an exception when its Visual is not Accepted (ignoring case and spaces); a blank Visual counts as an exception, not as accepted
# ASSUME: a row with a blank Weld Seam No is not a weld and is never an exception
# ASSUME: this tool owns the red fill on the log: every run removes the fill of rows that are not exceptions and fills the exception rows red
# ASSUME: an exception row with a blank Welder ID is copied but not counted for any welder; Welder IDs are trimmed and matched ignoring case; welders with equal counts come in no fixed order
# ASSUME: Weld Exceptions and Welder Exceptions are cleared and rebuilt on every run
GOAL     Highlight the weld log rows that are not Accepted, copy them to Weld Exceptions and count the exceptions of each welder, most first.
NAME     WeldFollowUp
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(log As Worksheet, exc As Worksheet, byWelder As Worksheet) As Long
SIG      modSteps.HighlightExceptions(log As Worksheet) As Long
SIG      modSteps.CopyExceptions(log As Worksheet, exc As Worksheet)
SIG      modSteps.CountByWelder(log As Worksheet, out As Worksheet)
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.HighlightExceptions
FLOW     modMain.RunHeadless -> modSteps.CopyExceptions
FLOW     modMain.RunHeadless -> modSteps.CountByWelder
FIXTURE  weld_01.xlsx sheet=Log cols=A:Weld Seam No,B:Welder ID,C:Visual rows=10 stale=4 data=A:text(6)+boundary("S-01","S-02","","S-04","S-05","S-06","S-07","S-08","S-09","S-10");B:text(4)+boundary("W-3","W-2","W-1","W-1"," w-1 ","W-3","","W-1","W-2","W-2");C:text(8)+boundary("Rejected","Repair","Rejected","Rejected","rejected"," accepted ","Rejected","","Accepted","Rejected")
FIXTURE  weld_01.xlsx sheet=Exc cols=A:Weld Seam No,B:Welder ID,C:Visual rows=0 stale=2
FIXTURE  weld_01.xlsx sheet=ByWelder cols=A:Welder ID,B:Exceptions rows=0 stale=2
TEST     fixture=weld_01.xlsx call=modMain.RunHeadless(Log, Exc, ByWelder) runs=2 assert=RESULT = 7 AND ERROR_COUNT = 0 AND FILL_COUNT(Log![Weld Seam No], #FF0000) = 7 AND FILL_COUNT(Log![Visual], #FF0000) = 7 AND COUNTA(Exc![Weld Seam No]) = 8 AND COUNTIF(Exc![Weld Seam No], "stale") = 0 AND COUNTA(ByWelder![Welder ID]) = 4 AND COUNTIF(ByWelder![Welder ID], "stale") = 0
TEST     fixture=weld_01.xlsx call=modMain.RunHeadless(Log, Exc, ByWelder) assert=RESULT = 7 AND COUNTA(Log![Weld Seam No]) = PRE_COUNTA(Log![Weld Seam No]) AND COUNTA(Log![Welder ID]) = PRE_COUNTA(Log![Welder ID]) AND COUNTA(Log![Visual]) = PRE_COUNTA(Log![Visual]) AND VALUE(Exc![Weld Seam No]2) = "S-01" AND VALUE(Exc![Visual]2) = "Rejected" AND VALUE(Exc![Weld Seam No]8) = "S-10" AND ROWS_USED(Exc) = 8 AND ROWS_USED(ByWelder) = 4 AND VALUE(ByWelder![Welder ID]2) = "W-1" AND VALUE(ByWelder![Exceptions]2) = 3 AND VALUE(ByWelder![Welder ID]3) = "W-2" AND VALUE(ByWelder![Exceptions]3) = 2 AND VALUE(ByWelder![Welder ID]4) = "W-3" AND VALUE(ByWelder![Exceptions]4) = 1
TEST     fixture=weld_01.xlsx call=modMain.RunHeadless(ByWelder, Exc, Log) assert=RESULT = -1 AND ERROR_COUNT = 1 AND FILL_COUNT(Log![Visual], #FF0000) = PRE_FILL_COUNT(Log![Visual], #FF0000)
---
VCL 0.6
TASK
GOAL     Start the weld follow-up tool on the workbook's own sheets and tell the user how many weld rows are exceptions.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "Weld Log"), SHEET(THIS_BOOK, "Weld Exceptions"), SHEET(THIS_BOOK, "Welder Exceptions"), into=n)
OP       MSG("Weld exceptions: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run the whole weld follow-up on the three sheets it is given and return how many weld rows are exceptions.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     log As Worksheet, exc As Worksheet, byWelder As Worksheet
RET      Long
VAR      flagged As Long
PRE      COL($log, "Weld Seam No") > 0; COL($log, "Welder ID") > 0; COL($log, "Visual") > 0
MODIFIES $log fills
MODIFIES $exc
MODIFIES $byWelder
OP       CALL(modSteps.HighlightExceptions, $log, into=flagged)
OP       CALL(modSteps.CopyExceptions, $log, $exc)
OP       CALL(modSteps.CountByWelder, $log, $byWelder)
OP       RETURN(flagged)
ERR      standard
---
VCL 0.6
TASK
GOAL     Fill red every weld log row whose Visual is not Accepted, remove the fill of every other row, and return how many rows are exceptions.
PROC     HighlightExceptions
KIND     Function
MODULE   modSteps
ARGS     log As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cVisual As Long
VAR      lastCol As Long
VAR      flagged As Long
PRE      COL($log, "Weld Seam No") > 0; COL($log, "Visual") > 0
MODIFIES $log fills
OP       LET(cSeam, COL($log, "Weld Seam No"))
OP       LET(cVisual, COL($log, "Visual"))
OP       LET(lastCol, LASTCOL($log, 1))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($log, cSeam) var=r
OP         IF IS_EMPTY(CELL($log, r, cSeam)) OR SAME(CELL($log, r, cVisual), "Accepted")
OP           FORMAT(ROW_RANGE($log, r, 1, lastCol), fill=none)
OP         ELSE
OP           FORMAT(ROW_RANGE($log, r, 1, lastCol), fill=red)
OP           LET(flagged, flagged + 1)
OP       RETURN(flagged)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Weld Exceptions sheet with the log's headings and a copy of every weld row whose Visual is not Accepted.
PROC     CopyExceptions
KIND     Sub
MODULE   modSteps
ARGS     log As Worksheet, exc As Worksheet
VAR      cSeam As Long
VAR      cVisual As Long
VAR      lastCol As Long
VAR      outRow As Long
PRE      COL($log, "Weld Seam No") > 0; COL($log, "Visual") > 0
MODIFIES $exc
OP       LET(cSeam, COL($log, "Weld Seam No"))
OP       LET(cVisual, COL($log, "Visual"))
OP       LET(lastCol, LASTCOL($log, 1))
OP       CLEAR(USED_RANGE($exc), what=all)
OP       COPY_ROW($log, 1, $exc, 1, cols=lastCol)
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($log, cSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($log, r, cSeam))
OP         SKIP_IF SAME(CELL($log, r, cVisual), "Accepted")
OP         COPY_ROW($log, r, $exc, outRow, cols=lastCol)
OP         LET(outRow, outRow + 1)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Welder Exceptions sheet with how many weld exceptions each welder has, the welder with the most first.
PROC     CountByWelder
KIND     Sub
MODULE   modSteps
ARGS     log As Worksheet, out As Worksheet
VAR      cSeam As Long
VAR      cWelder As Long
VAR      cVisual As Long
VAR      counts As Dictionary
VAR      got As Long
VAR      outRow As Long
PRE      COL($log, "Weld Seam No") > 0; COL($log, "Welder ID") > 0; COL($log, "Visual") > 0
MODIFIES $out
OP       LET(cSeam, COL($log, "Weld Seam No"))
OP       LET(cWelder, COL($log, "Welder ID"))
OP       LET(cVisual, COL($log, "Visual"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Welder ID", "Exceptions")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($log, cSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($log, r, cSeam))
OP         SKIP_IF SAME(CELL($log, r, cVisual), "Accepted")
OP         SKIP_IF IS_EMPTY(CELL($log, r, cWelder))
OP         DICT_ADD(counts, TRIM(CELL($log, r, cWelder)), 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         WRITE_ROW($out, outRow, k, got)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=2, order=desc, header=yes)
ERR      standard
