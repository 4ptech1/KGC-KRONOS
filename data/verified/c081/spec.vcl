VCL 0.6
TASK
# ASSUME: the Job Tracker headers in row 1 include "Customer", "Due Date" and "Status"; the real sheets are "Job Tracker" and "Customer Load", the TESTs use Jobs and Load
# ASSUME: the Customer Load sheet already exists; everything on it is cleared and rebuilt on every run, and it has no "Job ID" header
# ASSUME: a job is open when its Status is not Closed (ignoring case and extra spaces); a blank Status counts as open
# ASSUME: an open job is past due only when its Due Date is a real Excel date before today; due today, a blank or a text Due Date is not past due
# ASSUME: customer names are matched ignoring case and spaces at the ends, and shown as first typed; a row with a blank Customer is left out
# ASSUME: customers with the same number of open jobs come in no particular order
GOAL     Rebuild the Customer Load sheet with each customer's number of open jobs and how many of those are past their Due Date, most open jobs first.
PROC     BuildCustomerLoad
KIND     Function
MODULE   modCustomerLoad
ARGS     jobs As Worksheet, out As Worksheet
RET      Long
VAR      cCust As Long
VAR      cDue As Long
VAR      cStatus As Long
VAR      k As String
VAR      openJobs As Dictionary
VAR      lateJobs As Dictionary
VAR      nOpen As Long
VAR      nLate As Long
VAR      outRow As Long
PRE      COL($jobs, "Customer") > 0; COL($jobs, "Due Date") > 0; COL($jobs, "Status") > 0; COL($out, "Job ID") = 0
MODIFIES $out
OP       LET(cCust, COL($jobs, "Customer"))
OP       LET(cDue, COL($jobs, "Due Date"))
OP       LET(cStatus, COL($jobs, "Status"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Customer", "Open Jobs", "Past Due")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($jobs, cCust) var=r
OP         SKIP_IF IS_EMPTY(CELL($jobs, r, cCust))
OP         SKIP_IF SAME(CELL($jobs, r, cStatus), "Closed")
OP         LET(k, TRIM(CELL($jobs, r, cCust)))
OP         DICT_ADD(openJobs, k, 1, mode=sum)
OP         DICT_ADD(lateJobs, k, 0, mode=sum)
OP         IF IS_DATE(CELL($jobs, r, cDue))
OP           IF CELL($jobs, r, cDue) < TODAY THEN DICT_ADD(lateJobs, k, 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=cust in=KEYS(openJobs)
OP         DICT_GET(openJobs, cust, into=nOpen)
OP         DICT_GET(lateJobs, cust, into=nLate)
OP         WRITE_ROW($out, outRow, cust, nOpen, nLate)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=2, order=desc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off
ERR      standard
FIXTURE  load_01.xlsx sheet=Jobs cols=A:Job ID,B:Customer,C:Due Date,D:Status rows=15 stale=2 data=A:seq(1,1);B:values("Cedar Ltd"x2,"Birch Co"x2," birch co "x1,"Delta Fab"x4,"DELTA FAB"x1,"Zed Works"x2,"Acme"x2,""x1)+boundary("Cedar Ltd","Birch Co","Delta Fab","Zed Works","Acme","","Birch Co"," birch co ","Delta Fab","DELTA FAB","Delta Fab","Delta Fab","Acme","Cedar Ltd","Zed Works");C:reldates(-30,30)+boundary(1,-1,0,"",-10,-3,-7,5,-2,-4,9,"",4,-6,3);D:values("Open"x8,"Closed"x2,"In Progress"x1,"On Hold"x1," closed "x1,""x1,"CLOSED"x1)+boundary("Open","Open","Open","Open","Closed","Open","In Progress","On Hold"," closed ","Open","","Open","Closed","CLOSED","Open")
FIXTURE  load_01.xlsx sheet=Load cols=A:Customer,B:Open Jobs,C:Past Due rows=0 stale=6
FIXTURE  load_02.xlsx sheet=Jobs cols=A:Job ID,B:Customer,C:Due Date,D:Status rows=4 data=A:seq(1,1);B:values("Acme"x2,"Birch Co"x1,""x1)+boundary("Acme","Birch Co","","Acme");C:reldates(-10,10)+boundary(-1,0,-2,"");D:values("Closed"x3,"Open"x1)+boundary("Closed","Closed","Open","Closed")
FIXTURE  load_02.xlsx sheet=Load cols=A:Customer,B:Open Jobs,C:Past Due rows=0 stale=3
TEST     fixture=load_01.xlsx call=BuildCustomerLoad(Jobs, Load) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND ROWS_USED(Load) = 5 AND VALUE(Load!A1) = "Customer" AND VALUE(Load!B1) = "Open Jobs" AND VALUE(Load!C1) = "Past Due" AND VALUE(Load![Customer]2) = "Delta Fab" AND VALUE(Load![Open Jobs]2) = 4 AND VALUE(Load![Past Due]2) = 1 AND VALUE(Load![Customer]3) = "Birch Co" AND VALUE(Load![Open Jobs]3) = 3 AND VALUE(Load![Past Due]3) = 2 AND VALUE(Load![Customer]4) = "Zed Works" AND VALUE(Load![Open Jobs]4) = 2 AND VALUE(Load![Past Due]4) = 0 AND VALUE(Load![Customer]5) = "Cedar Ltd" AND VALUE(Load![Open Jobs]5) = 1 AND VALUE(Load![Past Due]5) = 0 AND COUNTIF(Load![Customer], "stale") = 0 AND COUNTIF(Load![Customer], "Acme") = 0 AND SUM(Load![Open Jobs]) = 10 AND SUM(Load![Past Due]) = 3
TEST     fixture=load_01.xlsx call=BuildCustomerLoad(Jobs, Load) assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(Jobs![Customer]) = PRE_COUNTA(Jobs![Customer]) AND COUNTA(Jobs![Due Date]) = PRE_COUNTA(Jobs![Due Date]) AND COUNTA(Jobs![Status]) = PRE_COUNTA(Jobs![Status]) AND COUNTIF(Jobs![Status], "Closed") = PRE_COUNTIF(Jobs![Status], "Closed") AND COUNTA(Jobs![Job ID]) = PRE_COUNTA(Jobs![Job ID])
TEST     fixture=load_02.xlsx call=BuildCustomerLoad(Jobs, Load) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND ROWS_USED(Load) = 1 AND VALUE(Load!A1) = "Customer" AND VALUE(Load!C1) = "Past Due" AND COUNTIF(Load![Customer], "stale") = 0 AND COUNTA(Jobs![Customer]) = PRE_COUNTA(Jobs![Customer])
TEST     fixture=load_01.xlsx call=BuildCustomerLoad(Load, Jobs) assert=RESULT = -1 AND ERROR_COUNT = 1
TEST     fixture=load_01.xlsx call=BuildCustomerLoad(Jobs, Jobs) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Jobs![Job ID]) = PRE_COUNTA(Jobs![Job ID])
