VCL 0.1
TASK
# ASSUME: Orders headers in row 1 include "Status" and "Order Date"
# ASSUME: "Cancelled" is matched ignoring case and extra spaces
# ASSUME: Order Date holds real dates; rows with a blank date sort to the bottom
PROC     RemoveCancelledAndSortByDate
KIND     Function
MODULE   modCleanup
ARGS     ws As Worksheet
RET      Long
VAR      cStatus As Long
VAR      cDate As Long
VAR      removed As Long
PRE      COL($ws, "Status") > 0; COL($ws, "Order Date") > 0
OP       LET(cStatus, COL($ws, "Status"))
OP       LET(cDate, COL($ws, "Order Date"))
OP       LET(removed, 0)
OP       LOOP_ROWS dir=up from=LASTROW($ws, cStatus) to=2 var=r
OP         IF UPPER(TRIM(CELL($ws, r, cStatus))) = "CANCELLED"
OP           DELETE_ROW($ws, r)
OP           LET(removed, removed + 1)
OP       SORT(USED_RANGE($ws), key=cDate, order=desc, header=yes)
OP       RETURN(removed)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  orders_01.xlsx sheet=Orders cols=A:Order No,B:Order Date,C:Customer,D:Status rows=40 data=A:seq(1,1);B:dates(2026-01-01,2026-09-30)+blank(2);C:text(12);D:values("Open"x20,"Cancelled"x8,"Closed"x12)
TEST     fixture=orders_01.xlsx call=RemoveCancelledAndSortByDate(Orders) runs=2 assert=COUNTIF(Orders![Status], "Cancelled") = 0 AND ERROR_COUNT = 0
TEST     fixture=orders_01.xlsx call=RemoveCancelledAndSortByDate(Orders) assert=RESULT >= PRE_COUNTIF(Orders![Status], "Cancelled") AND COUNTA(Orders![Status]) = PRE_COUNTA(Orders![Status]) - RESULT
TEST     fixture=orders_01.xlsx call=RemoveCancelledAndSortByDate(Orders) assert=VALUE(Orders![Order Date]2) >= VALUE(Orders![Order Date]3)
