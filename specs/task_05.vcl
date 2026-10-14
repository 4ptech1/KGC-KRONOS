VCL 0.1
TASK
# ASSUME: Weld Log headers (row 1) include "Weld Seam No", "Fit-up" and "Visual"; column order does not matter
# ASSUME: rows are NOT deleted; seams not yet accepted are copied to the Weld Log Exceptions sheet instead
# ASSUME: a seam is an exception when Fit-up or Visual is blank or anything other than "Accepted" (case and spaces ignored)
# ASSUME: the Weld Log Exceptions sheet is rebuilt every run with the Weld Log's header row
PROC     ListWeldLogExceptions
KIND     Function
MODULE   modWelding
ARGS     weldLog As Worksheet, out As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cFit As Long
VAR      cVisual As Long
VAR      lastCol As Long
VAR      outRow As Long
PRE      COL($weldLog, "Weld Seam No") > 0; COL($weldLog, "Fit-up") > 0; COL($weldLog, "Visual") > 0
OP       LET(cSeam, COL($weldLog, "Weld Seam No"))
OP       LET(cFit, COL($weldLog, "Fit-up"))
OP       LET(cVisual, COL($weldLog, "Visual"))
OP       LET(lastCol, LASTCOL($weldLog, 1))
OP       CLEAR(USED_RANGE($out), what=all)
OP       COPY_ROW($weldLog, 1, $out, 1, cols=lastCol)
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($weldLog, cSeam) var=r
OP         IF NOT IS_EMPTY(CELL($weldLog, r, cSeam))
OP           IF UPPER(TRIM(CELL($weldLog, r, cFit))) <> "ACCEPTED" OR UPPER(TRIM(CELL($weldLog, r, cVisual))) <> "ACCEPTED"
OP             COPY_ROW($weldLog, r, $out, outRow, cols=lastCol)
OP             LET(outRow, outRow + 1)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  weld_log_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Welder ID,C:Fit-up,D:Visual,E:Weld Date rows=50 data=A:seq(1,1);B:values("W-01"x25,"W-02"x25);C:values("Accepted"x44,"Rejected"x2,""x4);D:values("Accepted"x43,"Hold"x3,""x4);E:dates(2026-01-01,2026-09-30)
FIXTURE  weld_log_01.xlsx sheet=Exceptions cols=A:Weld Seam No rows=0
TEST     fixture=weld_log_01.xlsx call=ListWeldLogExceptions(WeldLog, Exceptions) runs=2 assert=RESULT >= 7 AND RESULT <= 13 AND COUNTA(Exceptions![Weld Seam No]) = RESULT + 1 AND ERROR_COUNT = 0
TEST     fixture=weld_log_01.xlsx call=ListWeldLogExceptions(WeldLog, Exceptions) assert=COUNTA(WeldLog![Weld Seam No]) = PRE_COUNTA(WeldLog![Weld Seam No]) AND COUNTA(WeldLog![Fit-up]) = PRE_COUNTA(WeldLog![Fit-up])
TEST     fixture=weld_log_01.xlsx call=ListWeldLogExceptions(Exceptions, Exceptions) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Exceptions![Weld Seam No]) = PRE_COUNTA(Exceptions![Weld Seam No])
