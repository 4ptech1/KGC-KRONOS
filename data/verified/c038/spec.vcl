VCL 0.6
TASK
# ASSUME: Dispatch headers in row 1 include "Dispatch No", "Customer", "Planned Date" and "Actual Date"
# ASSUME: the real sheets are "Dispatch" and "Customer Summary"; the summary sheet exists and is rebuilt on every run
# ASSUME: rows with a blank Customer are ignored; customers are matched ignoring case and extra spaces; the first spelling seen is written
# ASSUME: a dispatch is late only when both dates are real dates and Actual Date is after Planned Date; a blank or text date is never late
# ASSUME: the summary has the headings "Customer", "Dispatches" and "Late", and lists the customers from A to Z
GOAL     Count the dispatches per customer and how many of them were late, and write that on the Customer Summary sheet from A to Z.
PROC     SummariseDispatchesByCustomer
KIND     Function
MODULE   modDispatchSummary
ARGS     dispatch As Worksheet, out As Worksheet
RET      Long
VAR      cCust As Long
VAR      cPlan As Long
VAR      cAct As Long
VAR      total As Dictionary
VAR      late As Dictionary
VAR      nTotal As Long
VAR      nLate As Long
VAR      outRow As Long
PRE      COL($dispatch, "Dispatch No") > 0; COL($dispatch, "Customer") > 0; COL($dispatch, "Planned Date") > 0; COL($dispatch, "Actual Date") > 0
MODIFIES $out
OP       LET(cCust, COL($dispatch, "Customer"))
OP       LET(cPlan, COL($dispatch, "Planned Date"))
OP       LET(cAct, COL($dispatch, "Actual Date"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Customer", "Dispatches", "Late")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($dispatch, cCust) var=r
OP         SKIP_IF IS_EMPTY(CELL($dispatch, r, cCust))
OP         DICT_ADD(total, CLEAN_TEXT(CELL($dispatch, r, cCust)), 1, mode=sum)
OP         DICT_ADD(late, CLEAN_TEXT(CELL($dispatch, r, cCust)), 0, mode=sum)
OP         IF IS_DATE(CELL($dispatch, r, cPlan))
OP           IF IS_DATE(CELL($dispatch, r, cAct))
OP             IF CELL($dispatch, r, cAct) > CELL($dispatch, r, cPlan)
OP               DICT_ADD(late, CLEAN_TEXT(CELL($dispatch, r, cCust)), 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(total)
OP         DICT_GET(total, k, into=nTotal)
OP         DICT_GET(late, k, into=nLate)
OP         WRITE_ROW($out, outRow, k, nTotal, nLate)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       RETURN(COUNT(total))
PERF     screen_off calc_manual
ERR      standard
FIXTURE  dispatch_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Customer,C:Planned Date,D:Actual Date rows=12 data=A:seq(1,1);B:values("Zeta"x4,"Acme"x3," acme "x1,""x1,"Midway"x2,"Beta"x1)+boundary("Zeta","Zeta","Zeta","Zeta","Acme",""," acme ","Acme","Acme","Midway","Midway","Beta");C:dates(2026-03-01,2026-06-30)+boundary(2026-03-10,2026-03-10,2026-03-10,2026-03-10,2026-04-01,2026-04-01,"",2026-04-01,2026-05-01,2026-05-01,2026-05-02,2026-06-01);D:dates(2026-03-01,2026-06-30)+boundary(2026-03-10,2026-03-11,2026-03-09,"",2026-04-05,2026-04-09,2026-04-05,2026-04-02,2026-05-01,2026-05-20,2026-05-01,2026-06-01)
FIXTURE  dispatch_01.xlsx sheet=CustomerSummary cols=A:Customer,B:Dispatches,C:Late rows=0 stale=6
TEST     fixture=dispatch_01.xlsx call=SummariseDispatchesByCustomer(Dispatch, CustomerSummary) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND ROWS_USED(CustomerSummary) = 5 AND COUNTIF(CustomerSummary![Customer], "stale") = 0 AND VALUE(CustomerSummary![Customer]2) = "Acme" AND VALUE(CustomerSummary![Dispatches]2) = 4 AND VALUE(CustomerSummary![Late]2) = 2 AND VALUE(CustomerSummary![Customer]3) = "Beta" AND VALUE(CustomerSummary![Dispatches]3) = 1 AND VALUE(CustomerSummary![Late]3) = 0 AND VALUE(CustomerSummary![Customer]4) = "Midway" AND VALUE(CustomerSummary![Dispatches]4) = 2 AND VALUE(CustomerSummary![Late]4) = 1 AND VALUE(CustomerSummary![Customer]5) = "Zeta" AND VALUE(CustomerSummary![Dispatches]5) = 4 AND VALUE(CustomerSummary![Late]5) = 1
TEST     fixture=dispatch_01.xlsx call=SummariseDispatchesByCustomer(Dispatch, CustomerSummary) assert=SUM(CustomerSummary![Dispatches]) = 11 AND SUM(CustomerSummary![Late]) = 4 AND COUNTA(Dispatch![Dispatch No]) = PRE_COUNTA(Dispatch![Dispatch No]) AND COUNTA(Dispatch![Customer]) = PRE_COUNTA(Dispatch![Customer]) AND COUNTA(Dispatch![Actual Date]) = PRE_COUNTA(Dispatch![Actual Date]) AND ERROR_COUNT = 0
TEST     fixture=dispatch_01.xlsx call=SummariseDispatchesByCustomer(CustomerSummary, Dispatch) assert=RESULT = -1 AND ERROR_COUNT = 1
