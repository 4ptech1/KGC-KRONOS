VCL 0.6
TASK
# ASSUME: Issues headers in row 1 include "Issue No", "Batch No", "Qty Issued", "Job ID" and "Remark"; Jobs headers include "Job ID" and "Status"
# ASSUME: the Remark column already exists in the issue log; only that column changes
# ASSUME: Job IDs are matched ignoring case and extra spaces; Status is matched with SAME
# ASSUME: if a Job ID is listed more than once in the tracker, its first row decides; a blank Job ID is never a job
# ASSUME: an issue whose job is not in the tracker, or has a blank Job ID, is not for a closed job
GOAL     Write Closed job in the Remark column of every issue whose job is Closed in the Job Tracker, clear the remark on all other issues, and return how many issues are for closed jobs.
PROC     MarkClosedJobIssues
KIND     Function
MODULE   modConsumables
ARGS     issues As Worksheet, jobs As Worksheet
RET      Long
VAR      cIssueNo As Long
VAR      cIssueJob As Long
VAR      cRemark As Long
VAR      cJobId As Long
VAR      cJobStatus As Long
VAR      jobState As Dictionary
VAR      state As String
VAR      marked As Long
PRE      COL($issues, "Issue No") > 0; COL($issues, "Job ID") > 0; COL($issues, "Remark") > 0; COL($jobs, "Job ID") > 0; COL($jobs, "Status") > 0
MODIFIES $issues["Remark"]
OP       LET(cIssueNo, COL($issues, "Issue No"))
OP       LET(cIssueJob, COL($issues, "Job ID"))
OP       LET(cRemark, COL($issues, "Remark"))
OP       LET(cJobId, COL($jobs, "Job ID"))
OP       LET(cJobStatus, COL($jobs, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($jobs, cJobId) var=r
OP         SKIP_IF IS_EMPTY(CELL($jobs, r, cJobId))
OP         DICT_ADD(jobState, TRIM(CELL($jobs, r, cJobId)), TRIM(CELL($jobs, r, cJobStatus)), mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($issues, cIssueNo) var=i
OP         IF EXISTS(jobState, TRIM(CELL($issues, i, cIssueJob)))
OP           DICT_GET(jobState, TRIM(CELL($issues, i, cIssueJob)), into=state)
OP         ELSE
OP           LET(state, "")
OP         IF SAME(state, "Closed")
OP           SET_VALUE(CELL($issues, i, cRemark), "Closed job")
OP           LET(marked, marked + 1)
OP         ELSE
OP           CLEAR(CELL($issues, i, cRemark), what=contents)
OP       RETURN(marked)
ERR      standard
FIXTURE  consumables_01.xlsx sheet=Issues cols=A:Issue No,B:Batch No,C:Qty Issued,D:Job ID,E:Remark rows=8 data=A:seq(1,1);B:text(6);C:int(1,20);D:values("J-2"x2,"J-1"x1," j-3"x1,"J-4"x1,""x1,"J-9"x1,"J-5"x1)+boundary("J-2","J-1"," j-3","J-4","","J-9","J-5","J-2");E:values(""x3,"Closed job"x3,"old"x2)+boundary("","Closed job","Closed job","old","Closed job","","old","")
FIXTURE  consumables_01.xlsx sheet=Jobs cols=A:Job ID,B:Status rows=8 data=A:values("J-1"x1,"J-2"x1," J-3 "x1,"J-4"x2,""x1,"J-5"x1,"J-6"x1)+boundary("J-1","J-2"," J-3 ","J-4","J-4","","J-5","J-6");B:values("Open"x3,"Closed"x4,"closed "x1)+boundary("Open","Closed","closed ","Open","Closed","Closed","Closed","Open")
TEST     fixture=consumables_01.xlsx call=MarkClosedJobIssues(Issues, Jobs) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(Issues![Remark], "Closed job") = 4 AND COUNTA(Issues![Remark]) = 5
TEST     fixture=consumables_01.xlsx call=MarkClosedJobIssues(Issues, Jobs) assert=RESULT = 4 AND ERROR_COUNT = 0 AND VALUE(Issues![Remark]2) = "Closed job" AND VALUE(Issues![Remark]4) = "Closed job" AND VALUE(Issues![Remark]8) = "Closed job" AND VALUE(Issues![Remark]9) = "Closed job" AND COUNTIF(Issues![Remark], "old") = 0
TEST     fixture=consumables_01.xlsx call=MarkClosedJobIssues(Issues, Jobs) assert=COUNTA(Issues![Issue No]) = PRE_COUNTA(Issues![Issue No]) AND SUM(Issues![Qty Issued]) = PRE_SUM(Issues![Qty Issued]) AND COUNTA(Jobs![Job ID]) = PRE_COUNTA(Jobs![Job ID])
TEST     fixture=consumables_01.xlsx call=MarkClosedJobIssues(Jobs, Issues) assert=RESULT = -1 AND ERROR_COUNT = 1
