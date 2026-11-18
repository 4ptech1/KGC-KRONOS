VCL 0.6
TASK
# ASSUME: Job Tracker headers in row 1 are Job ID, Customer, Status and Dispatch Flag; Dispatch headers are Dispatch No, Job ID, Planned Date and Actual Date
# ASSUME: the three output sheets Ready To Ship and Customer Dispatches exist already; they are rebuilt from scratch on every run (real sheets: Job Tracker, Dispatch, Ready To Ship, Customer Dispatches; TESTs use Jobs, Dispatch, ReadyShip, CustDisp)
# ASSUME: a Dispatch row has an Actual Date when that cell is not empty; a job with such a row is Dispatched whatever its Status
# ASSUME: Job IDs are matched after trimming and ignoring case; Status Complete is matched with SAME; customers are matched the same way and shown as first typed
# ASSUME: a Tracker row with a blank Job ID is skipped and left as it is; a customer with no Dispatched jobs is listed with 0; a blank Customer is not listed
# ASSUME: Ready To Ship and Customer Dispatches are sorted A to Z as text; Ready To Ship holds trimmed values and Customer Dispatches has the headings Customer and Dispatched Jobs
GOAL     Mark each job Dispatched, Ready or blank in the Job Tracker, list the Ready jobs and count Dispatched jobs per customer on two rebuilt sheets, and tell the user how many jobs are Ready.
PROC     UpdateDispatchReadiness
KIND     Function
MODULE   modDispatch
ARGS     jobs As Worksheet, disp As Worksheet, ready As Worksheet, cust As Worksheet
RET      Long
VAR      cJob As Long
VAR      cCust As Long
VAR      cStatus As Long
VAR      cFlag As Long
VAR      dJob As Long
VAR      dAct As Long
VAR      shipped As Dictionary
VAR      counts As Dictionary
VAR      nReady As Long
VAR      rr As Long
VAR      cr As Long
VAR      got As Long
PRE      COL($jobs, "Job ID") > 0; COL($jobs, "Customer") > 0; COL($jobs, "Status") > 0; COL($jobs, "Dispatch Flag") > 0; COL($disp, "Job ID") > 0; COL($disp, "Actual Date") > 0
MODIFIES $jobs["Dispatch Flag"]
MODIFIES $ready
MODIFIES $cust
OP       LET(cJob, COL($jobs, "Job ID"))
OP       LET(cCust, COL($jobs, "Customer"))
OP       LET(cStatus, COL($jobs, "Status"))
OP       LET(cFlag, COL($jobs, "Dispatch Flag"))
OP       LET(dJob, COL($disp, "Job ID"))
OP       LET(dAct, COL($disp, "Actual Date"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($disp, dJob) var=d
OP         SKIP_IF IS_EMPTY(CELL($disp, d, dJob))
OP         IF NOT IS_EMPTY(CELL($disp, d, dAct)) THEN DICT_ADD(shipped, TRIM(CELL($disp, d, dJob)), 1, mode=keep)
OP       CLEAR(USED_RANGE($ready), what=all)
OP       CLEAR(USED_RANGE($cust), what=all)
OP       WRITE_ROW($ready, 1, "Job ID", "Customer", "Status", "Dispatch Flag")
OP       WRITE_ROW($cust, 1, "Customer", "Dispatched Jobs")
OP       LET(rr, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($jobs, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($jobs, r, cJob))
OP         IF NOT IS_EMPTY(CELL($jobs, r, cCust)) THEN DICT_ADD(counts, TRIM(CELL($jobs, r, cCust)), 0, mode=sum)
OP         IF EXISTS(shipped, TRIM(CELL($jobs, r, cJob)))
OP           SET_VALUE(CELL($jobs, r, cFlag), "Dispatched")
OP           IF NOT IS_EMPTY(CELL($jobs, r, cCust)) THEN DICT_ADD(counts, TRIM(CELL($jobs, r, cCust)), 1, mode=sum)
OP         ELSE
OP           IF SAME(CELL($jobs, r, cStatus), "Complete")
OP             SET_VALUE(CELL($jobs, r, cFlag), "Ready")
OP             WRITE_ROW($ready, rr, TRIM(CELL($jobs, r, cJob)), TRIM(CELL($jobs, r, cCust)), TRIM(CELL($jobs, r, cStatus)), "Ready")
OP             LET(rr, rr + 1)
OP             LET(nReady, nReady + 1)
OP           ELSE
OP             CLEAR(CELL($jobs, r, cFlag), what=contents)
OP       LET(cr, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         WRITE_ROW($cust, cr, k, got)
OP         LET(cr, cr + 1)
OP       IF rr > 2 THEN SORT(USED_RANGE($ready), key=1, order=asc, header=yes)
OP       IF cr > 2 THEN SORT(USED_RANGE($cust), key=1, order=asc, header=yes)
OP       MSG("Jobs ready to ship: " & nReady)
OP       RETURN(nReady)
PERF     screen_off
ERR      standard
FIXTURE  dispatch_01.xlsx sheet=Jobs cols=A:Job ID,B:Customer,C:Status,D:Dispatch Flag rows=10 data=A:values("J-01"x1," j-02 "x1,"J-08"x1,"J-04 "x1,"j-05"x1,"J-06"x1,""x1,"J-03"x1,"J-09"x1," J-10"x1)+boundary("J-01"," j-02 ","J-08","J-04 ","j-05","J-06","","J-03","J-09"," J-10");B:values("Acme"x2,"acme "x1,""x1,"Beta"x2,"Gamma"x2,"Delta"x2)+boundary("Acme","acme ","","Beta","Acme","Gamma","Delta","Beta","Gamma","Delta");C:values("Complete"x6," Complete"x1,"Open"x2,"In Progress"x1)+boundary("Complete","Complete","Complete"," Complete","Open","In Progress","Complete","Complete","Open","Complete");D:values("old"x4,""x5,"Ready"x1)+boundary("old","","","old","","old","old","","Ready","")
FIXTURE  dispatch_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Job ID,C:Planned Date,D:Actual Date rows=8 data=A:seq(1,1);B:values("J-03"x1,"J-01"x2,"J-02 "x1,"j-05"x1," J-10"x1,""x1,"J-99"x1)+boundary("J-03","J-01","J-02 ","j-05"," J-10","","J-01","J-99");C:dates(2026-03-01,2026-03-31);D:dates(2026-03-01,2026-03-31)+boundary("","")
FIXTURE  dispatch_01.xlsx sheet=ReadyShip cols=A:Job ID,B:Customer,C:Status,D:Dispatch Flag rows=0 stale=2
FIXTURE  dispatch_01.xlsx sheet=CustDisp cols=A:Customer,B:Dispatched Jobs rows=0 stale=2
TEST     fixture=dispatch_01.xlsx call=UpdateDispatchReadiness(Jobs, Dispatch, ReadyShip, CustDisp) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Jobs![Dispatch Flag], "Dispatched") = 4 AND COUNTIF(Jobs![Dispatch Flag], "Ready") = 3 AND COUNTIF(Jobs![Dispatch Flag], "old") = 1 AND COUNTA(Jobs![Dispatch Flag]) = 9 AND COUNTA(ReadyShip![Job ID]) = 4 AND COUNTIF(ReadyShip![Job ID], "stale") = 0 AND VALUE(ReadyShip![Job ID]2) = "J-03" AND VALUE(ReadyShip![Job ID]3) = "J-04" AND VALUE(ReadyShip![Job ID]4) = "J-08" AND VALUE(ReadyShip![Customer]2) = "Beta" AND VALUE(ReadyShip![Dispatch Flag]4) = "Ready"
TEST     fixture=dispatch_01.xlsx call=UpdateDispatchReadiness(Jobs, Dispatch, ReadyShip, CustDisp) assert=ERROR_COUNT = 0 AND COUNTA(CustDisp![Customer]) = 5 AND COUNTIF(CustDisp![Customer], "stale") = 0 AND VALUE(CustDisp![Customer]2) = "Acme" AND VALUE(CustDisp![Dispatched Jobs]2) = 3 AND VALUE(CustDisp![Customer]3) = "Beta" AND VALUE(CustDisp![Dispatched Jobs]3) = 0 AND VALUE(CustDisp![Customer]4) = "Delta" AND VALUE(CustDisp![Dispatched Jobs]4) = 1 AND VALUE(CustDisp![Customer]5) = "Gamma" AND VALUE(CustDisp![Dispatched Jobs]5) = 0 AND COUNTA(Jobs![Job ID]) = PRE_COUNTA(Jobs![Job ID]) AND COUNTA(Dispatch![Job ID]) = PRE_COUNTA(Dispatch![Job ID])
TEST     fixture=dispatch_01.xlsx call=UpdateDispatchReadiness(Dispatch, Jobs, ReadyShip, CustDisp) assert=RESULT = -1 AND ERROR_COUNT = 1
