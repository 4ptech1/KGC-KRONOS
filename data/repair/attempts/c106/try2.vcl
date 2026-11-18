VCL 0.6
TASK
# ASSUME: Job Tracker headers in row 1 are Job ID, Customer, Status, Dispatch Flag; Dispatch headers are Dispatch No, Job ID, Planned Date, Actual Date
# ASSUME: Actual Date holds real dates; a Dispatch row with a blank or non-date Actual Date does not count as dispatched
# ASSUME: Job IDs are matched trimmed and ignoring case; a Tracker row with a blank Job ID is left untouched and not listed
# ASSUME: Status Complete is matched ignoring case and spaces; Job ID, Customer and Status are written trimmed on the output sheets
# ASSUME: Ready To Ship and Customer Dispatches are existing sheets that are rebuilt on every run; customers with no dispatched job are listed with 0
# ASSUME: the sort is on the trimmed text, ignoring case
GOAL     Mark each job Dispatched, Ready or blank in the Job Tracker, list the Ready jobs and the Dispatched count per customer on two rebuilt sheets, and return how many jobs are Ready.
PROC     MarkDispatchReadiness
KIND     Function
MODULE   modDispatch
ARGS     jobs As Worksheet, disp As Worksheet, ship As Worksheet, summ As Worksheet
RET      Long
VAR      cJob As Long
VAR      cCust As Long
VAR      cStatus As Long
VAR      cFlag As Long
VAR      dJob As Long
VAR      dDate As Long
VAR      sent As Dictionary
VAR      counts As Dictionary
VAR      readyN As Long
VAR      outRow As Long
VAR      got As Long
PRE      COL($jobs, "Job ID") > 0; COL($jobs, "Customer") > 0; COL($jobs, "Status") > 0; COL($jobs, "Dispatch Flag") > 0; COL($disp, "Job ID") > 0; COL($disp, "Actual Date") > 0
MODIFIES $jobs["Dispatch Flag"]
MODIFIES $ship
MODIFIES $summ
OP       LET(cJob, COL($jobs, "Job ID"))
OP       LET(cCust, COL($jobs, "Customer"))
OP       LET(cStatus, COL($jobs, "Status"))
OP       LET(cFlag, COL($jobs, "Dispatch Flag"))
OP       LET(dJob, COL($disp, "Job ID"))
OP       LET(dDate, COL($disp, "Actual Date"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($disp, dJob) var=d
OP         SKIP_IF IS_EMPTY(CELL($disp, d, dJob))
OP         IF IS_DATE(CELL($disp, d, dDate)) THEN DICT_ADD(sent, TRIM(CELL($disp, d, dJob)), 1, mode=keep)
OP       CLEAR(USED_RANGE($ship), what=all)
OP       WRITE_ROW($ship, 1, "Job ID", "Customer", "Status", "Dispatch Flag")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($jobs, cJob) var=j
OP         SKIP_IF IS_EMPTY(CELL($jobs, j, cJob))
OP         IF NOT IS_EMPTY(CELL($jobs, j, cCust)) THEN DICT_ADD(counts, TRIM(CELL($jobs, j, cCust)), IIF(EXISTS(sent, TRIM(CELL($jobs, j, cJob))), 1, 0), mode=sum)
OP         IF EXISTS(sent, TRIM(CELL($jobs, j, cJob)))
OP           SET_VALUE(CELL($jobs, j, cFlag), "Dispatched")
OP         ELSE
OP           IF SAME(CELL($jobs, j, cStatus), "Complete")
OP             SET_VALUE(CELL($jobs, j, cFlag), "Ready")
OP             LET(readyN, readyN + 1)
OP             WRITE_ROW($ship, readyN + 1, TRIM(CELL($jobs, j, cJob)), TRIM(CELL($jobs, j, cCust)), TRIM(CELL($jobs, j, cStatus)), "Ready")
OP           ELSE
OP             CLEAR(CELL($jobs, j, cFlag), what=contents)
OP       SORT(USED_RANGE($ship), key=1, order=asc, header=yes)
OP       CLEAR(USED_RANGE($summ), what=all)
OP       WRITE_ROW($summ, 1, "Customer", "Dispatched")
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         WRITE_ROW($summ, outRow, k, got)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($summ), key=1, order=asc, header=yes)
OP       MSG("Jobs ready: " & readyN)
OP       RETURN(readyN)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  disp_01.xlsx sheet=Jobs cols=A:Job ID,B:Customer,C:Status,D:Dispatch Flag rows=10 data=A:text(8)+boundary("J-01"," j-02 ","J-03","j-04","J-05 ","","J-07","B-08","J-09","J-10");B:text(8)+boundary("Acme","Beta","Acme","Cora","Beta","Zed","Cora","Acme","Dora","Beta");C:text(8)+boundary("Complete","Complete","Open","Complete","In Progress","Complete"," complete ","Complete","Open","Complete");D:values("old"x10)
FIXTURE  disp_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Job ID,C:Planned Date,D:Actual Date rows=7 data=A:seq(1,1);B:text(8)+boundary("J-02","j-04 "," j-01","J-03","J-10","j-10 ","");C:dates(2026-01-01,2026-01-31);D:dates(2026-02-01,2026-02-28)+boundary("","")
FIXTURE  disp_01.xlsx sheet=ReadyShip cols=A:Job ID,B:Customer,C:Status,D:Dispatch Flag rows=0 stale=3
FIXTURE  disp_01.xlsx sheet=CustDisp cols=A:Customer,B:Dispatched rows=0 stale=3
TEST     fixture=disp_01.xlsx call=MarkDispatchReadiness(Jobs, Dispatch, ReadyShip, CustDisp) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(Jobs![Dispatch Flag], "Dispatched") = 3 AND COUNTIF(Jobs![Dispatch Flag], "Ready") = 4 AND COUNTIF(Jobs![Dispatch Flag], "old") = 1 AND COUNTA(Jobs![Dispatch Flag]) = 9 AND ROWS_USED(ReadyShip) = 5 AND COUNTIF(ReadyShip![Job ID], "stale") = 0
TEST     fixture=disp_01.xlsx call=MarkDispatchReadiness(Jobs, Dispatch, ReadyShip, CustDisp) assert=VALUE(ReadyShip![Job ID]2) = "B-08" AND VALUE(ReadyShip![Job ID]3) = "j-02" AND VALUE(ReadyShip![Job ID]4) = "j-04" AND VALUE(ReadyShip![Job ID]5) = "J-07" AND VALUE(ReadyShip![Customer]3) = "Beta" AND VALUE(ReadyShip![Status]5) = "complete" AND COUNTIF(ReadyShip![Dispatch Flag], "Ready") = 4
TEST     fixture=disp_01.xlsx call=MarkDispatchReadiness(Jobs, Dispatch, ReadyShip, CustDisp) assert=ROWS_USED(CustDisp) = 5 AND VALUE(CustDisp![Customer]2) = "Acme" AND VALUE(CustDisp![Dispatched]2) = 2 AND VALUE(CustDisp![Customer]3) = "Beta" AND VALUE(CustDisp![Dispatched]3) = 1 AND VALUE(CustDisp![Dispatched]4) = 0 AND VALUE(CustDisp![Customer]5) = "Dora" AND VALUE(CustDisp![Dispatched]5) = 0 AND COUNTIF(CustDisp![Customer], "stale") = 0
TEST     fixture=disp_01.xlsx call=MarkDispatchReadiness(Jobs, Dispatch, ReadyShip, CustDisp) assert=COUNTA(Jobs![Job ID]) = PRE_COUNTA(Jobs![Job ID]) AND COUNTA(Dispatch![Job ID]) = PRE_COUNTA(Dispatch![Job ID]) AND COUNTA(Dispatch![Actual Date]) = PRE_COUNTA(Dispatch![Actual Date])
TEST     fixture=disp_01.xlsx call=MarkDispatchReadiness(Dispatch, Jobs, ReadyShip, CustDisp) assert=RESULT = -1 AND ERROR_COUNT = 1
