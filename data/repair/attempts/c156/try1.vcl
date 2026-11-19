VCL 0.6
TASK
# ASSUME: Dispatch headers in row 1 include "Customer", "Planned Date" and "Actual Date"; the real sheets are Dispatch and On Time, and On Time already exists
# ASSUME: customers are grouped ignoring case and stray spaces, and shown as first typed; rows with a blank Customer are ignored
# ASSUME: a dispatch counts as delivered only when its Actual Date is a real date; it is late only when its Planned Date is also a real date and the Actual Date is after it
# ASSUME: a customer with no delivered dispatch still gets a row with 0 and 0; the function returns the number of customers listed
GOAL     Rebuild the On Time sheet with each customer, its dispatches that have an actual date and how many of those were late, customers A to Z.
PROC     BuildOnTimeReport
KIND     Function
MODULE   modDispatch
ARGS     disp As Worksheet, out As Worksheet
RET      Long
VAR      cCust As Long
VAR      cPlan As Long
VAR      cAct As Long
VAR      done As Dictionary
VAR      late As Dictionary
VAR      doneN As Long
VAR      lateN As Long
VAR      outRow As Long
PRE      COL($disp, "Customer") > 0; COL($disp, "Planned Date") > 0; COL($disp, "Actual Date") > 0
MODIFIES $out
OP       LET(cCust, COL($disp, "Customer"))
OP       LET(cPlan, COL($disp, "Planned Date"))
OP       LET(cAct, COL($disp, "Actual Date"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Customer", "Dispatched", "Late")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($disp, cCust) var=r
OP         SKIP_IF IS_EMPTY(CELL($disp, r, cCust))
OP         DICT_ADD(done, TRIM(CELL($disp, r, cCust)), 0, mode=sum)
OP         DICT_ADD(late, TRIM(CELL($disp, r, cCust)), 0, mode=sum)
OP         IF IS_DATE(CELL($disp, r, cAct))
OP           DICT_ADD(done, TRIM(CELL($disp, r, cCust)), 1, mode=sum)
OP           IF IS_DATE(CELL($disp, r, cPlan))
OP             IF CELL($disp, r, cAct) > CELL($disp, r, cPlan)
OP               DICT_ADD(late, TRIM(CELL($disp, r, cCust)), 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(done)
OP         DICT_GET(done, k, into=doneN)
OP         DICT_GET(late, k, into=lateN)
OP         WRITE_ROW($out, outRow, k, doneN, lateN)
OP         LET(outRow, outRow + 1)
OP       IF outRow > 3 THEN SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       RETURN(outRow - 2)
ERR      standard
FIXTURE  disp_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Customer,C:Planned Date,D:Actual Date rows=12 data=A:seq(1,1);B:values("Zeta Ltd"x3,"zeta ltd"x2,"Alpha Co"x3,"ALPHA CO "x1,"Mid Works"x2,""x1)+boundary("Zeta Ltd","zeta ltd","Alpha Co","ALPHA CO ","","Mid Works","Mid Works","Zeta Ltd","Alpha Co","Zeta Ltd","Alpha Co","zeta ltd");C:dates(2026-03-01,2026-03-31)+boundary("2026-03-10","2026-03-10","2026-03-10","2026-03-10","2026-03-10","2026-03-10","2026-03-10","2026-03-10");D:dates(2026-03-01,2026-04-15)+blank(2)+boundary("2026-03-10","2026-03-11","2026-03-09","2026-03-12","2026-03-20","","2026-03-10","2026-03-11")
FIXTURE  disp_01.xlsx sheet=OnTime cols=A:Customer,B:Dispatched,C:Late rows=0 stale=3
TEST     fixture=disp_01.xlsx call=BuildOnTimeReport(Dispatch, OnTime) runs=2 assert=ERROR_COUNT = 0 AND COUNTIF(OnTime![Customer], "stale") = 0 AND COUNTA(OnTime![Customer]) = 4
TEST     fixture=disp_01.xlsx call=BuildOnTimeReport(Dispatch, OnTime) assert=RESULT = 3 AND COUNTA(Dispatch![Dispatch No]) = PRE_COUNTA(Dispatch![Dispatch No])
TEST     fixture=disp_01.xlsx call=BuildOnTimeReport(OnTime, Dispatch) assert=RESULT = -1 AND ERROR_COUNT = 1
