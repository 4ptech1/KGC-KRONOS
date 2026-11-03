VCL 0.6
TASK
# ASSUME: Dispatch headers in row 1 include "Job ID" and "Actual Date"; the Job Tracker headers include "Job ID"; "Dispatch Remark" is added after the last heading when it is missing
# ASSUME: Job IDs are compared after trimming spaces at both ends, ignoring case, so J-101 and " j-101 " are the same job
# ASSUME: a dispatch counts only when its Actual Date is a real date; a blank or text entry such as TBC does not
# ASSUME: a dispatch row with no Job ID is ignored; a job whose dispatch rows have no Actual Date is not dispatched
# ASSUME: the last Job Tracker row is found from Job ID, each job is listed once there, and every job that is not dispatched has its remark cleared
GOAL     Write Dispatched in the Job Tracker's Dispatch Remark column for each job with at least one dispatch that has an Actual Date, clear it on the other jobs, and return how many jobs are dispatched.
PROC     MarkDispatchedJobs
KIND     Function
MODULE   modDispatch
ARGS     dispatch As Worksheet, tracker As Worksheet
RET      Long
VAR      cJobD As Long
VAR      cActual As Long
VAR      cJobT As Long
VAR      cRemark As Long
VAR      sent As Dictionary
VAR      n As Long
PRE      COL($dispatch, "Job ID") > 0; COL($dispatch, "Actual Date") > 0; COL($tracker, "Job ID") > 0
MODIFIES $tracker
OP       LET(cJobD, COL($dispatch, "Job ID"))
OP       LET(cActual, COL($dispatch, "Actual Date"))
OP       LET(cJobT, COL($tracker, "Job ID"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($dispatch, cJobD) var=r
OP         SKIP_IF IS_EMPTY(CELL($dispatch, r, cJobD))
OP         IF IS_DATE(CELL($dispatch, r, cActual))
OP           DICT_ADD(sent, TRIM(CELL($dispatch, r, cJobD)), 1, mode=keep)
OP       IF COL($tracker, "Dispatch Remark") = 0
OP         SET_VALUE(CELL($tracker, 1, LASTCOL($tracker, 1) + 1), "Dispatch Remark")
OP       LET(cRemark, COL($tracker, "Dispatch Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($tracker, cJobT) var=t
OP         CLEAR(CELL($tracker, t, cRemark), what=contents)
OP         IF EXISTS(sent, TRIM(CELL($tracker, t, cJobT)))
OP           SET_VALUE(CELL($tracker, t, cRemark), "Dispatched")
OP           LET(n, n + 1)
OP       RETURN(n)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  dispatch_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Job ID,C:Actual Date rows=12 data=A:seq(1,1);B:values("J-101"x1,"J-101 "x1," J-102"x1,"J-103"x1,"J-105"x2," J-106 "x1,""x1,"J-107"x1,"J-108"x1,"J-999"x1,"J-102"x1)+boundary("J-101","J-101 "," J-102","J-103","J-105","J-105"," J-106 ","","J-107","J-108","J-999","J-102");C:reldates(-90,-1)+boundary(-10,-9,"",-20,"",-30,-31,-8,"",-5,-3,"")
FIXTURE  dispatch_01.xlsx sheet=JobTracker cols=A:Job ID,B:Customer,C:Status,D:Dispatch Remark rows=9 data=A:values("J-101"x1,"J-102"x1," J-103"x1,"J-104"x1,""x1,"J-105 "x1,"J-106"x1,"J-107"x1,"J-108"x1)+boundary("J-101","J-102"," J-103","J-104","","J-105 ","J-106","J-107","J-108");B:text(8);C:values("Open"x5,"Closed"x4);D:values("Dispatched"x4,"Hold"x2,""x3)+boundary("","Dispatched","Dispatched","Hold","Dispatched","","Hold","","Dispatched")
FIXTURE  dispatch_02.xlsx sheet=Dispatch cols=A:Dispatch No,B:Job ID,C:Actual Date rows=3 data=A:seq(1,1);B:values("J-201"x1,"J-202"x1,"J-203"x1)+boundary("J-201","J-202","J-203");C:reldates(-30,-1)+boundary(-4,"",-2)
FIXTURE  dispatch_02.xlsx sheet=JobTracker cols=A:Job ID,B:Customer,C:Status rows=4 data=A:values("J-201"x1,"J-202"x1,"J-203"x1,"J-204"x1)+boundary("J-201","J-202","J-203","J-204");B:text(8);C:values("Open"x3,""x1)+boundary("","Open","Open","Open")
FIXTURE  dispatch_03.xlsx sheet=Dispatch cols=A:Dispatch No,B:Job ID,C:Actual Date rows=3 data=A:seq(1,1);B:values("J-301"x2,"J-302"x1);C:values("TBC"x2,"pending"x1)
FIXTURE  dispatch_03.xlsx sheet=JobTracker cols=A:Job ID,B:Customer,C:Status,D:Dispatch Remark rows=2 data=A:values("J-301"x1,"J-302"x1)+boundary("J-301","J-302");B:text(8);C:values("Open"x2);D:values("Dispatched"x1,""x1)+boundary("Dispatched","")
TEST     fixture=dispatch_01.xlsx call=MarkDispatchedJobs(Dispatch, JobTracker) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTIF(JobTracker![Dispatch Remark], "Dispatched") = 5 AND COUNTA(JobTracker![Dispatch Remark]) = 6 AND VALUE(JobTracker![Dispatch Remark]2) = "Dispatched" AND VALUE(JobTracker![Dispatch Remark]4) = "Dispatched" AND VALUE(JobTracker![Dispatch Remark]7) = "Dispatched" AND VALUE(JobTracker![Dispatch Remark]8) = "Dispatched" AND VALUE(JobTracker![Dispatch Remark]10) = "Dispatched"
TEST     fixture=dispatch_01.xlsx call=MarkDispatchedJobs(Dispatch, JobTracker) assert=RESULT = 5 AND COUNTA(JobTracker![Job ID]) = PRE_COUNTA(JobTracker![Job ID]) AND COUNTA(JobTracker![Customer]) = PRE_COUNTA(JobTracker![Customer]) AND COUNTA(JobTracker![Status]) = PRE_COUNTA(JobTracker![Status]) AND COUNTA(Dispatch![Job ID]) = PRE_COUNTA(Dispatch![Job ID])
TEST     fixture=dispatch_02.xlsx call=MarkDispatchedJobs(Dispatch, JobTracker) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(JobTracker!D1) = "Dispatch Remark" AND COUNTIF(JobTracker![Dispatch Remark], "Dispatched") = 2 AND COUNTA(JobTracker![Dispatch Remark]) = 3 AND VALUE(JobTracker![Dispatch Remark]2) = "Dispatched" AND VALUE(JobTracker![Dispatch Remark]4) = "Dispatched" AND COUNTA(JobTracker!E:E) = 0
TEST     fixture=dispatch_03.xlsx call=MarkDispatchedJobs(Dispatch, JobTracker) assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTA(JobTracker![Dispatch Remark]) = 1
TEST     fixture=dispatch_01.xlsx call=MarkDispatchedJobs(JobTracker, Dispatch) assert=RESULT = -1 AND ERROR_COUNT = 1
