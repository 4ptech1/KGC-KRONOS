VCL 0.6
TASK
# ASSUME: NDE Report headers in row 1 include "Weld Seam No" and "Result"
# ASSUME: every row from 2 down to the last Weld Seam No is checked, across all columns up to the last header
# ASSUME: Fail and Pass are matched ignoring case and extra spaces; any other Result, or a blank one, gets no fill
# ASSUME: old fills are cleared on every data row before the new colour is set; the header row is left alone
GOAL     Fill each NDE report row red when its Result is Fail and green when it is Pass, clear the fill on other rows and tell how many rows failed.
PROC     ColourNdeRows
KIND     Function
MODULE   modNde
ARGS     ws As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cResult As Long
VAR      cLast As Long
VAR      failed As Long
PRE      COL($ws, "Weld Seam No") > 0; COL($ws, "Result") > 0
MODIFIES $ws fills
OP       LET(cSeam, COL($ws, "Weld Seam No"))
OP       LET(cResult, COL($ws, "Result"))
OP       LET(cLast, LASTCOL($ws, 1))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cSeam) var=r
OP         FORMAT(ROW_RANGE($ws, r, 1, cLast), fill=none)
OP         IF SAME(CELL($ws, r, cResult), "Fail")
OP           FORMAT(ROW_RANGE($ws, r, 1, cLast), fill=red)
OP           LET(failed, failed + 1)
OP         ELSE
OP           IF SAME(CELL($ws, r, cResult), "Pass")
OP             FORMAT(ROW_RANGE($ws, r, 1, cLast), fill=green)
OP       MSG("NDE rows failed: " & failed)
OP       RETURN(failed)
PERF     screen_off
ERR      standard
FIXTURE  nde_01.xlsx sheet=NdeReport cols=A:Weld Seam No,B:Method,C:Result,D:Report Date rows=12 stale=3 data=A:seq(1,1);B:values("RT"x4,"UT"x4,"MT"x2,"PT"x2);C:values("Fail"x3,"FAIL "x1,"Pass"x4,"pass"x1,"Repair"x1,""x1,"Re-test"x1)+boundary("Fail","Pass","FAIL ","Repair","pass","","Pass","Fail","Re-test","Pass","Fail","Pass");D:dates(2026-01-05,2026-09-30)+boundary("")
FIXTURE  nde_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=nde_01.xlsx call=ColourNdeRows(NdeReport) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND FILL_COUNT(NdeReport![Result], #FF0000) = 4 AND FILL_COUNT(NdeReport![Result], #00FF00) = 5 AND FILL_COUNT(NdeReport![Weld Seam No], #FF0000) = 4 AND FILL_COUNT(NdeReport![Report Date], #FF0000) = 4 AND FILL_COUNT(NdeReport![Method], #00FF00) = 5 AND FILL_COUNT(NdeReport![Report Date], #00FF00) = 5
TEST     fixture=nde_01.xlsx call=ColourNdeRows(NdeReport) assert=FILL_COUNT(NdeReport![Weld Seam No]2, #FF0000) = 1 AND FILL_COUNT(NdeReport![Weld Seam No]3, #00FF00) = 1 AND FILL_COUNT(NdeReport![Weld Seam No]4, #FF0000) = 1 AND FILL_COUNT(NdeReport![Weld Seam No]5, #FF0000) = 0 AND FILL_COUNT(NdeReport![Weld Seam No]5, #00FF00) = 0 AND FILL_COUNT(NdeReport![Weld Seam No]6, #00FF00) = 1 AND FILL_COUNT(NdeReport![Weld Seam No]7, #FF0000) = 0 AND FILL_COUNT(NdeReport![Weld Seam No]7, #00FF00) = 0 AND FILL_COUNT(NdeReport![Weld Seam No]9, #00FF00) = 0 AND FILL_COUNT(NdeReport![Weld Seam No]10, #00FF00) = 0 AND COUNTA(NdeReport![Weld Seam No]) = PRE_COUNTA(NdeReport![Weld Seam No]) AND COUNTA(NdeReport![Result]) = PRE_COUNTA(NdeReport![Result]) AND COUNTA(NdeReport![Report Date]) = PRE_COUNTA(NdeReport![Report Date])
TEST     fixture=nde_01.xlsx call=ColourNdeRows(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
