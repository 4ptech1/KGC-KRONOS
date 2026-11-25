VCL 0.7
BUILD
# ASSUME: the real sheets are "Job Tracker", "NCR Register", "Jobs On Hold" and "Unmatched NCRs"; the TESTs use Jobs, Ncr, Hold and Unmatched
# ASSUME: the two output sheets exist and are rebuilt on every run; their headings are written by the tool
# ASSUME: Job IDs are trimmed and matched ignoring case; an NCR counts as Open when its Status is Open (SAME)
# ASSUME: a job is on hold when its Status is not Closed and it has at least one Open NCR; ties keep no set order
# ASSUME: an NCR row with a blank NCR No is skipped; an NCR with a blank or unknown Job ID is listed as unmatched
GOAL     Count the open NCRs of each job, list the jobs on hold with most NCRs first, list the NCRs whose job is unknown, and report the number of jobs on hold.
NAME     JobQualityTool
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(jobs As Worksheet, ncr As Worksheet, hold As Worksheet, unmatched As Worksheet) As Long
SIG      modSteps.CountOpenNcrs(jobs As Worksheet, ncr As Worksheet)
SIG      modSteps.ListJobsOnHold(jobs As Worksheet, hold As Worksheet) As Long
SIG      modSteps.ListUnmatchedNcrs(jobs As Worksheet, ncr As Worksheet, unmatched As Worksheet)
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.CountOpenNcrs
FLOW     modMain.RunHeadless -> modSteps.ListJobsOnHold
FLOW     modMain.RunHeadless -> modSteps.ListUnmatchedNcrs
FIXTURE  jq_01.xlsx sheet=Jobs cols=A:Job ID,B:Customer,C:Status,D:Open NCRs rows=6 data=A:values(" J-1"x1,"J-2"x1,""x1,"j-3 "x1,"J-4"x1,"J-5"x1)+boundary(" J-1","J-2","","j-3 ","J-4","J-5");B:text(6);C:values("Open"x5,"Closed"x1)+boundary("Open","Closed","Open","Open","Open","Open");D:values("old"x6)
FIXTURE  jq_01.xlsx sheet=Ncr cols=A:NCR No,B:Job ID,C:Status rows=10 data=A:seq(1,1)+boundary(1,2,3,4,"",6,7,8,9,10);B:values(" j-1"x1,"J-1"x2,"J-2 "x1,"J-8"x1,"J-3"x2,"j-3 "x1,""x1,"J-9"x1)+boundary(" j-1","J-1","J-1","J-2 ","J-8","J-3","j-3 ","J-3","","J-9");C:values("Open"x9,"Closed"x1)+boundary("Open","Open","Closed","Open","Open","Open","Open","Open","Open","Open")
FIXTURE  jq_01.xlsx sheet=Hold cols=A:Job ID,B:Customer,C:Status,D:Open NCRs rows=0 stale=2
FIXTURE  jq_01.xlsx sheet=Unmatched cols=A:NCR No,B:Job ID,C:Status rows=0 stale=2
TEST     fixture=jq_01.xlsx call=modMain.RunHeadless(Jobs, Ncr, Hold, Unmatched) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(Jobs![Open NCRs]2) = 2 AND VALUE(Jobs![Open NCRs]3) = 1 AND VALUE(Jobs![Open NCRs]4) = "old" AND VALUE(Jobs![Open NCRs]5) = 3 AND VALUE(Jobs![Open NCRs]6) = 0 AND VALUE(Jobs![Open NCRs]7) = 0 AND ROWS_USED(Hold) = 3 AND VALUE(Hold![Open NCRs]2) = 3 AND VALUE(Hold![Open NCRs]3) = 2 AND COUNTIF(Hold![Job ID], "stale") = 0 AND ROWS_USED(Unmatched) = 3 AND COUNTIF(Unmatched![Job ID], "J-9") = 1 AND COUNTIF(Unmatched![Status], "stale") = 0
TEST     fixture=jq_01.xlsx call=modMain.RunHeadless(Jobs, Ncr, Hold, Unmatched) assert=RESULT = 2 AND COUNTA(Jobs![Job ID]) = PRE_COUNTA(Jobs![Job ID]) AND COUNTA(Ncr![NCR No]) = PRE_COUNTA(Ncr![NCR No]) AND COUNTA(Ncr![Job ID]) = PRE_COUNTA(Ncr![Job ID]) AND VALUE(Hold![Job ID]1) = "Job ID" AND VALUE(Unmatched![NCR No]1) = "NCR No"
TEST     fixture=jq_01.xlsx call=modMain.RunHeadless(Ncr, Jobs, Hold, Unmatched) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.7
TASK
GOAL     Start the job quality tool on the workbook's own sheets and tell the user how many jobs are on hold.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "Job Tracker"), SHEET(THIS_BOOK, "NCR Register"), SHEET(THIS_BOOK, "Jobs On Hold"), SHEET(THIS_BOOK, "Unmatched NCRs"), into=n)
OP       MSG("Jobs on hold: " & n)
ERR      standard
---
VCL 0.7
TASK
GOAL     Run all three steps of the job quality tool on the four sheets it is given and return how many jobs are on hold.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     jobs As Worksheet, ncr As Worksheet, hold As Worksheet, unmatched As Worksheet
RET      Long
VAR      onHold As Long
PRE      COL($jobs, "Job ID") > 0; COL($jobs, "Customer") > 0; COL($jobs, "Status") > 0; COL($jobs, "Open NCRs") > 0; COL($ncr, "NCR No") > 0; COL($ncr, "Job ID") > 0; COL($ncr, "Status") > 0
MODIFIES $jobs["Open NCRs"]
MODIFIES $hold
MODIFIES $unmatched
OP       CALL(modSteps.CountOpenNcrs, $jobs, $ncr)
OP       CALL(modSteps.ListJobsOnHold, $jobs, $hold, into=onHold)
OP       CALL(modSteps.ListUnmatchedNcrs, $jobs, $ncr, $unmatched)
OP       RETURN(onHold)
ERR      standard
---
VCL 0.7
TASK
GOAL     Write into each job's Open NCRs column how many Open NCRs the NCR register has for it, 0 when none.
PROC     CountOpenNcrs
KIND     Sub
MODULE   modSteps
ARGS     jobs As Worksheet, ncr As Worksheet
VAR      cJob As Long
VAR      cOpen As Long
VAR      nJob As Long
VAR      nStatus As Long
VAR      counts As Dictionary
VAR      got As Long
PRE      COL($jobs, "Job ID") > 0; COL($jobs, "Customer") > 0; COL($jobs, "Status") > 0; COL($jobs, "Open NCRs") > 0; COL($ncr, "NCR No") > 0; COL($ncr, "Job ID") > 0; COL($ncr, "Status") > 0
MODIFIES $jobs["Open NCRs"]
OP       LET(nJob, COL($ncr, "Job ID"))
OP       LET(nStatus, COL($ncr, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, nJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($ncr, r, nJob))
OP         IF SAME(CELL($ncr, r, nStatus), "Open")
OP           DICT_ADD(counts, TRIM(CELL($ncr, r, nJob)), 1, mode=sum)
OP       LET(cJob, COL($jobs, "Job ID"))
OP       LET(cOpen, COL($jobs, "Open NCRs"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($jobs, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($jobs, r, cJob))
OP         DICT_GET(counts, TRIM(CELL($jobs, r, cJob)), into=got)
OP         SET_VALUE(CELL($jobs, r, cOpen), got)
ERR      standard
---
VCL 0.7
TASK
GOAL     Rebuild the Jobs On Hold sheet with the jobs that are not Closed and have an open NCR, the most NCRs first.
PROC     ListJobsOnHold
KIND     Function
MODULE   modSteps
ARGS     jobs As Worksheet, hold As Worksheet
RET      Long
VAR      cJob As Long
VAR      cCust As Long
VAR      cStatus As Long
VAR      cOpen As Long
VAR      outRow As Long
PRE      COL($jobs, "Job ID") > 0; COL($jobs, "Customer") > 0; COL($jobs, "Status") > 0; COL($jobs, "Open NCRs") > 0
MODIFIES $hold
OP       LET(cJob, COL($jobs, "Job ID"))
OP       LET(cCust, COL($jobs, "Customer"))
OP       LET(cStatus, COL($jobs, "Status"))
OP       LET(cOpen, COL($jobs, "Open NCRs"))
OP       CLEAR(USED_RANGE($hold), what=all)
OP       WRITE_ROW($hold, 1, "Job ID", "Customer", "Status", "Open NCRs")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($jobs, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($jobs, r, cJob))
OP         SKIP_IF SAME(CELL($jobs, r, cStatus), "Closed")
OP         SKIP_IF NOT IS_NUMBER(CELL($jobs, r, cOpen))
OP         SKIP_IF CELL($jobs, r, cOpen) < 1
OP         WRITE_ROW($hold, outRow, CELL($jobs, r, cJob), CELL($jobs, r, cCust), CELL($jobs, r, cStatus), CELL($jobs, r, cOpen))
OP         LET(outRow, outRow + 1)
OP       IF outRow > 2
OP         SORT(USED_RANGE($hold), key=4, order=desc, header=yes)
OP       RETURN(outRow - 2)
ERR      standard
---
VCL 0.7
TASK
GOAL     Rebuild the Unmatched NCRs sheet with every NCR whose Job ID is not in the job tracker.
PROC     ListUnmatchedNcrs
KIND     Sub
MODULE   modSteps
ARGS     jobs As Worksheet, ncr As Worksheet, unmatched As Worksheet
VAR      cJob As Long
VAR      nNo As Long
VAR      nJob As Long
VAR      nStatus As Long
VAR      ids As Dictionary
VAR      outRow As Long
PRE      COL($jobs, "Job ID") > 0; COL($ncr, "NCR No") > 0; COL($ncr, "Job ID") > 0; COL($ncr, "Status") > 0
MODIFIES $unmatched
OP       LET(cJob, COL($jobs, "Job ID"))
OP       LET(nNo, COL($ncr, "NCR No"))
OP       LET(nJob, COL($ncr, "Job ID"))
OP       LET(nStatus, COL($ncr, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($jobs, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($jobs, r, cJob))
OP         DICT_ADD(ids, TRIM(CELL($jobs, r, cJob)), 1, mode=keep)
OP       CLEAR(USED_RANGE($unmatched), what=all)
OP       WRITE_ROW($unmatched, 1, "NCR No", "Job ID", "Status")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, nNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($ncr, r, nNo))
OP         SKIP_IF EXISTS(ids, TRIM(CELL($ncr, r, nJob)))
OP         WRITE_ROW($unmatched, outRow, CELL($ncr, r, nNo), CELL($ncr, r, nJob), CELL($ncr, r, nStatus))
OP         LET(outRow, outRow + 1)
ERR      standard
