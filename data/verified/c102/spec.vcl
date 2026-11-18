VCL 0.6
TASK
# ASSUME: Dispatch headers in row 1 include "Dispatch No", "Planned Date", "Actual Date" and "Remark"
# ASSUME: Late means Actual Date is after Planned Date; Overdue means no Actual Date and Planned Date before today
# ASSUME: all other rows with a usable Planned Date get a cleared Remark; a row whose dates are not real dates keeps its Remark
GOAL     Mark each dispatch Late or Overdue in the Remark column, clear the Remark on the others and return how many are Late or Overdue together.
PROC     MarkLateDispatches
KIND     Function
MODULE   modDispatch
ARGS     ws As Worksheet
RET      Long
VAR      cNo As Long
VAR      cPlan As Long
VAR      cAct As Long
VAR      cRemark As Long
VAR      n As Long
PRE      COL($ws, "Dispatch No") > 0; COL($ws, "Planned Date") > 0; COL($ws, "Actual Date") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Remark"]
OP       LET(cNo, COL($ws, "Dispatch No"))
OP       LET(cPlan, COL($ws, "Planned Date"))
OP       LET(cAct, COL($ws, "Actual Date"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cNo) var=r
OP         IF IS_DATE(CELL($ws, r, cAct))
OP           IF IS_DATE(CELL($ws, r, cPlan))
OP             IF CELL($ws, r, cAct) > CELL($ws, r, cPlan)
OP               SET_VALUE(CELL($ws, r, cRemark), "Late")
OP               LET(n, n + 1)
OP             ELSE
OP               CLEAR(CELL($ws, r, cRemark), what=contents)
OP         ELSE
OP           IF IS_EMPTY(CELL($ws, r, cAct))
OP             IF IS_DATE(CELL($ws, r, cPlan))
OP               IF CELL($ws, r, cPlan) < TODAY
OP                 SET_VALUE(CELL($ws, r, cRemark), "Overdue")
OP                 LET(n, n + 1)
OP               ELSE
OP                 CLEAR(CELL($ws, r, cRemark), what=contents)
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  dispatch_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Customer,C:Planned Date,D:Actual Date,E:Remark rows=10 data=A:seq(1,1);B:text(8);C:reldates(-30,30)+boundary(-5,-5,-5,-1,0,3,-2,5,"","");D:reldates(-30,30)+boundary(-3,-5,-7,"","","",0,4,"",2);E:values("old"x10)
FIXTURE  dispatch_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=dispatch_01.xlsx call=MarkLateDispatches(Dispatch) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Dispatch![Remark], "Late") = 2 AND COUNTIF(Dispatch![Remark], "Overdue") = 1 AND COUNTIF(Dispatch![Remark], "old") = 2 AND COUNTA(Dispatch![Remark]) = 6
TEST     fixture=dispatch_01.xlsx call=MarkLateDispatches(Dispatch) assert=VALUE(Dispatch![Remark]2) = "Late" AND VALUE(Dispatch![Remark]3) = "" AND VALUE(Dispatch![Remark]4) = "" AND VALUE(Dispatch![Remark]5) = "Overdue" AND VALUE(Dispatch![Remark]6) = "" AND VALUE(Dispatch![Remark]7) = "" AND VALUE(Dispatch![Remark]8) = "Late" AND VALUE(Dispatch![Remark]9) = "" AND VALUE(Dispatch![Remark]10) = "old" AND VALUE(Dispatch![Remark]11) = "old"
TEST     fixture=dispatch_01.xlsx call=MarkLateDispatches(Dispatch) assert=COUNTA(Dispatch![Dispatch No]) = PRE_COUNTA(Dispatch![Dispatch No]) AND COUNTA(Dispatch![Planned Date]) = PRE_COUNTA(Dispatch![Planned Date]) AND COUNTA(Dispatch![Actual Date]) = PRE_COUNTA(Dispatch![Actual Date]) AND COUNTA(Dispatch![Customer]) = PRE_COUNTA(Dispatch![Customer]) AND ERROR_COUNT = 0
TEST     fixture=dispatch_01.xlsx call=MarkLateDispatches(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
