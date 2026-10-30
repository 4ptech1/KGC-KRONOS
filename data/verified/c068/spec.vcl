VCL 0.6
TASK
# ASSUME: Dispatch headers in row 1 include "Dispatch No" and "Actual Date"
# ASSUME: D-205 is matched ignoring case and extra spaces; if it appears twice, only the first row is dated
# ASSUME: today's date replaces anything already in that row's Actual Date cell
GOAL     Find dispatch D-205 on the Dispatch sheet, write today's date in its Actual Date cell and return its row number, or 0 if it is not there.
PROC     StampD205ActualDate
KIND     Function
MODULE   modDispatch
ARGS     ws As Worksheet
RET      Long
VAR      cNo As Long
VAR      cActual As Long
PRE      COL($ws, "Dispatch No") > 0; COL($ws, "Actual Date") > 0
MODIFIES $ws["Actual Date"]
OP       LET(cNo, COL($ws, "Dispatch No"))
OP       LET(cActual, COL($ws, "Actual Date"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cNo) var=r
OP         IF SAME(CELL($ws, r, cNo), "D-205")
OP           SET_VALUE(CELL($ws, r, cActual), TODAY)
OP           RETURN(r)
OP       RETURN(0)
ERR      standard
FIXTURE  dispatch_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Customer,C:Planned Date,D:Actual Date rows=8 data=A:values(" d-205 "x1,"D-202"x1,""x1,"D-206"x1,"D-205"x1,"D-207"x1,"D-203"x1,"D-208"x1)+boundary(" d-205 ","D-202","","D-206","D-205");B:text(10);C:reldates(-20,20)+boundary(0,-3,5,-6,2);D:reldates(-30,-1)+boundary("",-2,"",-4,"")
FIXTURE  dispatch_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  dispatch_02.xlsx sheet=Dispatch cols=A:Dispatch No,B:Customer,C:Planned Date,D:Actual Date rows=4 data=A:values("D-201"x1,"D-204"x1,"D-206"x1,"D-205"x1)+boundary("D-201","D-204","D-206","D-205");B:text(10);C:reldates(-20,20)+boundary(-1,3,-2,0);D:reldates(-30,-1)+boundary(-3,"",-8,-6)
FIXTURE  dispatch_03.xlsx sheet=Dispatch cols=A:Dispatch No,B:Customer,C:Planned Date,D:Actual Date rows=5 data=A:values("D-201"x1,"D-2050"x1,"D-204"x1,"D 205"x1,"D-206"x1);B:text(10);C:reldates(-20,20);D:reldates(-30,-1)+boundary("")
TEST     fixture=dispatch_01.xlsx call=StampD205ActualDate(Dispatch) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(Dispatch![Actual Date]2) = VALUE(Dispatch![Planned Date]2) AND COUNTA(Dispatch![Actual Date]) = PRE_COUNTA(Dispatch![Actual Date]) + 1
TEST     fixture=dispatch_02.xlsx call=StampD205ActualDate(Dispatch) assert=RESULT = 5 AND ERROR_COUNT = 0 AND VALUE(Dispatch![Actual Date]5) = VALUE(Dispatch![Planned Date]5) AND COUNTA(Dispatch![Actual Date]) = PRE_COUNTA(Dispatch![Actual Date])
TEST     fixture=dispatch_03.xlsx call=StampD205ActualDate(Dispatch) assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTA(Dispatch![Actual Date]) = PRE_COUNTA(Dispatch![Actual Date])
TEST     fixture=dispatch_01.xlsx call=StampD205ActualDate(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
