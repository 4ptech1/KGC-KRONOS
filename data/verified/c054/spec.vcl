VCL 0.6
TASK
# ASSUME: Dispatch headers in row 1 include "Dispatch No", "Planned Date" and "Actual Date"; Customer is not used
# ASSUME: the whole row means every column from 1 to the last heading, and the data runs down to the last Dispatch No
# ASSUME: there is an Actual Date when that cell is not blank, whatever it holds; a cell with only spaces counts as blank
# ASSUME: a row is red only when its Actual Date is blank and its Planned Date is a real date before today; a Planned Date typed as text or left blank is never red
# ASSUME: old fills are cleared row by row from row 2 down, so the heading row keeps its fill; the red count is the Function's result
GOAL     Colour each Dispatch row green when it has an Actual Date, or red when it has none and its Planned Date is before today, and report how many rows are red.
PROC     HighlightDispatchRows
KIND     Function
MODULE   modDispatch
ARGS     ws As Worksheet
RET      Long
VAR      cNo As Long
VAR      cPlan As Long
VAR      cActual As Long
VAR      lastCol As Long
VAR      redRows As Long
PRE      COL($ws, "Dispatch No") > 0; COL($ws, "Planned Date") > 0; COL($ws, "Actual Date") > 0
MODIFIES $ws fills
OP       LET(cNo, COL($ws, "Dispatch No"))
OP       LET(cPlan, COL($ws, "Planned Date"))
OP       LET(cActual, COL($ws, "Actual Date"))
OP       LET(lastCol, LASTCOL($ws, 1))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cNo) var=r
OP         FORMAT(ROW_RANGE($ws, r, 1, lastCol), fill=none)
OP         IF NOT IS_EMPTY(CELL($ws, r, cActual))
OP           FORMAT(ROW_RANGE($ws, r, 1, lastCol), fill=green)
OP         ELSE
OP           IF IS_DATE(CELL($ws, r, cPlan))
OP             IF CELL($ws, r, cPlan) < TODAY
OP               FORMAT(ROW_RANGE($ws, r, 1, lastCol), fill=red)
OP               LET(redRows, redRows + 1)
OP       RETURN(redRows)
PERF     screen_off
ERR      standard
FIXTURE  dispatch_rows_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Customer,C:Planned Date,D:Actual Date rows=12 stale=6 data=A:seq(1001,1);B:text(8);C:reldates(-30,30)+boundary(-1,0,1,-10,5,-3,"","",-20,10,3,-2);D:reldates(-30,0)+boundary("","","",-5,-1,"","",0,"","",0,-2)
FIXTURE  dispatch_rows_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=dispatch_rows_01.xlsx call=HighlightDispatchRows(Dispatch) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND FILL_COUNT(Dispatch![Customer], #FF0000) = 3 AND FILL_COUNT(Dispatch![Customer], #00FF00) = 5
TEST     fixture=dispatch_rows_01.xlsx call=HighlightDispatchRows(Dispatch) assert=FILL_COUNT(Dispatch![Actual Date], #FF0000) = 3 AND FILL_COUNT(Dispatch![Dispatch No], #00FF00) = 5 AND FILL_COUNT(Dispatch![Customer]2, #FF0000) = 1 AND FILL_COUNT(Dispatch![Customer]3, #FF0000) = 0 AND COUNTA(Dispatch![Dispatch No]) = PRE_COUNTA(Dispatch![Dispatch No]) AND COUNTA(Dispatch![Actual Date]) = PRE_COUNTA(Dispatch![Actual Date])
TEST     fixture=dispatch_rows_01.xlsx call=HighlightDispatchRows(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
