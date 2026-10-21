VCL 0.6
TASK
# ASSUME: Dispatch headers in row 1 include "Dispatch No", "Planned Date", "Actual Date" and "Status"
# ASSUME: no Actual Date means the cell is empty or only spaces
# ASSUME: Late needs a real Planned Date before today; On time needs real Actual and Planned Dates
# ASSUME: a row delivered after its Planned Date, or without real dates, keeps its Status
# ASSUME: a row with a blank Dispatch No is skipped
GOAL     Mark each undelivered dispatch past its planned date as Late and each dispatch delivered by its planned date as On time, and return how many are Late.
PROC     MarkDispatchStatus
KIND     Function
MODULE   modDispatch
ARGS     ws As Worksheet
RET      Long
VAR      cNo As Long
VAR      cPlanned As Long
VAR      cActual As Long
VAR      cStatus As Long
VAR      late As Long
PRE      COL($ws, "Dispatch No") > 0; COL($ws, "Planned Date") > 0; COL($ws, "Actual Date") > 0; COL($ws, "Status") > 0
MODIFIES $ws["Status"]
OP       LET(cNo, COL($ws, "Dispatch No"))
OP       LET(cPlanned, COL($ws, "Planned Date"))
OP       LET(cActual, COL($ws, "Actual Date"))
OP       LET(cStatus, COL($ws, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cNo))
OP         IF IS_EMPTY(CELL($ws, r, cActual))
OP           IF IS_DATE(CELL($ws, r, cPlanned))
OP             IF CELL($ws, r, cPlanned) < TODAY
OP               SET_VALUE(CELL($ws, r, cStatus), "Late")
OP               LET(late, late + 1)
OP         ELSE
OP           IF IS_DATE(CELL($ws, r, cActual))
OP             IF IS_DATE(CELL($ws, r, cPlanned))
OP               IF CELL($ws, r, cActual) <= CELL($ws, r, cPlanned)
OP                 SET_VALUE(CELL($ws, r, cStatus), "On time")
OP       RETURN(late)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  dispatch_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Customer,C:Planned Date,D:Actual Date,E:Status rows=10 data=A:text(5)+boundary("D-01","D-02","D-03","D-04","D-05","D-06","D-07","","D-08","D-09");B:text(8);C:reldates(-30,30)+boundary(-5,-1,0,5,-10,-10,-10,-20,"","");D:reldates(-30,30)+boundary("","","","",-10,-11,-9,"","",-3);E:values("Pending"x9,"Late"x1)+boundary("Pending","Pending","Pending","Pending","Late","Pending","Pending","Pending","Pending","Pending")
FIXTURE  dispatch_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=dispatch_01.xlsx call=MarkDispatchStatus(Dispatch) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND COUNTIF(Dispatch![Status], "Late") = 2 AND COUNTIF(Dispatch![Status], "On time") = 2
TEST     fixture=dispatch_01.xlsx call=MarkDispatchStatus(Dispatch) assert=RESULT = 2 AND COUNTIF(Dispatch![Status], "Pending") = 6 AND COUNTA(Dispatch![Status]) = PRE_COUNTA(Dispatch![Status]) AND VALUE(Dispatch![Status]2) = "Late" AND VALUE(Dispatch![Status]3) = "Late" AND VALUE(Dispatch![Status]4) = "Pending" AND VALUE(Dispatch![Status]5) = "Pending" AND VALUE(Dispatch![Status]6) = "On time" AND VALUE(Dispatch![Status]7) = "On time" AND VALUE(Dispatch![Status]8) = "Pending" AND VALUE(Dispatch![Status]9) = "Pending"
TEST     fixture=dispatch_01.xlsx call=MarkDispatchStatus(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
