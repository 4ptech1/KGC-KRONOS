VCL 0.6
TASK
# ASSUME: Dispatch headers in row 1 include "Dispatch No", "Planned Date" and "Actual Date"
# ASSUME: every row from 2 down to the last Dispatch No is checked
# ASSUME: already past means the Planned Date is before today; a Planned Date of today is not past
# ASSUME: an Actual Date holding only spaces counts as blank
# ASSUME: a Planned Date that is not a real date cannot be past, so that row gets no fill
# ASSUME: only the Dispatch No cells are filled or cleared; no other cell changes
GOAL     Highlight the Dispatch No cell yellow when the Actual Date is blank and the Planned Date is past, clear the fill on the other Dispatch No cells and tell how many are highlighted.
PROC     HighlightLateDispatches
KIND     Function
MODULE   modDispatch
ARGS     ws As Worksheet
RET      Long
VAR      cNo As Long
VAR      cPlanned As Long
VAR      cActual As Long
VAR      n As Long
PRE      COL($ws, "Dispatch No") > 0; COL($ws, "Planned Date") > 0; COL($ws, "Actual Date") > 0
MODIFIES $ws fills
OP       LET(cNo, COL($ws, "Dispatch No"))
OP       LET(cPlanned, COL($ws, "Planned Date"))
OP       LET(cActual, COL($ws, "Actual Date"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cNo) var=r
OP         FORMAT(CELL($ws, r, cNo), fill=none)
OP         IF IS_EMPTY(CELL($ws, r, cActual))
OP           IF IS_DATE(CELL($ws, r, cPlanned))
OP             IF CELL($ws, r, cPlanned) < TODAY
OP               FORMAT(CELL($ws, r, cNo), fill=yellow)
OP               LET(n, n + 1)
OP       MSG("Dispatches highlighted: " & n)
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  dispatch_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Customer,C:Planned Date,D:Actual Date rows=12 stale=3 data=A:seq(5001,1);B:values("Alpha"x4,"Beta"x4,"Gamma"x4);C:reldates(-30,30)+boundary(-1,0,1,-1,-10,"",-30,-2,-4,20,-3,-7);D:reldates(-40,-1)+boundary("","","",-2,"","",-5,""," ",-4,-9,"")
FIXTURE  dispatch_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=dispatch_01.xlsx call=HighlightLateDispatches(Dispatch) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND FILL_COUNT(Dispatch![Dispatch No], #FFFF00) = 5 AND FILL_COUNT(Dispatch![Dispatch No], #FF0000) = 0
TEST     fixture=dispatch_01.xlsx call=HighlightLateDispatches(Dispatch) assert=FILL_COUNT(Dispatch![Dispatch No]2, #FFFF00) = 1 AND FILL_COUNT(Dispatch![Dispatch No]6, #FFFF00) = 1 AND FILL_COUNT(Dispatch![Dispatch No]9, #FFFF00) = 1 AND FILL_COUNT(Dispatch![Dispatch No]10, #FFFF00) = 1 AND FILL_COUNT(Dispatch![Dispatch No]13, #FFFF00) = 1 AND FILL_COUNT(Dispatch![Dispatch No]3, #FFFF00) = 0 AND FILL_COUNT(Dispatch![Dispatch No]4, #FFFF00) = 0 AND FILL_COUNT(Dispatch![Dispatch No]5, #FFFF00) = 0 AND FILL_COUNT(Dispatch![Dispatch No]7, #FFFF00) = 0 AND FILL_COUNT(Dispatch![Dispatch No]8, #FFFF00) = 0 AND FILL_COUNT(Dispatch![Customer], #FF0000) = PRE_FILL_COUNT(Dispatch![Customer], #FF0000) AND COUNTA(Dispatch![Dispatch No]) = PRE_COUNTA(Dispatch![Dispatch No]) AND COUNTA(Dispatch![Actual Date]) = PRE_COUNTA(Dispatch![Actual Date])
TEST     fixture=dispatch_01.xlsx call=HighlightLateDispatches(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
