VCL 0.6
BUILD
# ASSUME: the real sheets are "Job Tracker", "NCR Register", "Jobs On Hold" and "Unmatched NCRs"; the TESTs use Jobs, Ncr, Hold and Unmatched
# ASSUME: Jobs On Hold and Unmatched NCRs already exist in the workbook and are rebuilt on every run
# ASSUME: Job Tracker headers include "Job ID", "Customer", "Status" and "Open NCRs"; NCR Register headers include "NCR No", "Job ID" and "Status"
# ASSUME: Job IDs are matched ignoring case and spaces at the ends; an NCR counts as open when its Status is Open, ignoring case and spaces
# ASSUME: a job is on hold when its Status is not Closed (a blank Status is not Closed) and its Open NCRs is above 0
# ASSUME: Job Tracker rows with a blank Job ID are left alone; an NCR with a blank Job ID is listed as unmatched; rows with a blank NCR No are ignored
# ASSUME: the output sheets copy every column of the source sheet's header row
GOAL     Count each job's Open NCRs into the Job Tracker, list jobs on hold and NCRs with no matching job on two rebuilt sheets, and report how many jobs are on hold.
NAME     JobQualityTool
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(jobs As Worksheet, ncr As Worksheet, hold As Worksheet, unmatched As Worksheet) As Long
SIG      modSteps.CountOpenNcrs(jobs As Worksheet, ncr As Worksheet)
SIG      modSteps.WriteOnHold(jobs As Worksheet, hold As Worksheet) As Long
SIG      modSteps.WriteUnmatched(jobs As Worksheet, ncr As Worksheet, out As Worksheet)
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.CountOpenNcrs
FLOW     modMain.RunHeadless -> modSteps.WriteOnHold
FLOW     modMain.RunHeadless -> modSteps.WriteUnmatched
FIXTURE  jobq_01.xlsx sheet=Jobs cols=A:Job ID,B:Customer,C:Status,D:Open NCRs rows=8 data=A:values("J-1"x1,"J-2"x1,"J-3"x1,""x1," J-4 "x1,"J-5"x1,"J-6"x1,"J-7"x1)+boundary("J-1","J-2","J-3",""," J-4 ","J-5","J-6","J-7");B:text(6);C:values("Active"x5,"Closed"x1," closed "x1,""x1)+boundary("Active","Closed","Active","Active"," closed ","Active","","Active");D:int(8,9)+boundary(9,9,9,9,9,9,9,9)
FIXTURE  jobq_01.xlsx sheet=Ncr cols=A:NCR No,B:Job ID,C:Status rows=14 data=A:values("N-1"x1,"N-2"x1,"N-3"x1,"N-4"x1,"N-5"x1,"N-6"x1,"N-7"x1,"N-8"x1,"N-9"x1,"N-10"x1,""x1,"N-11"x1,"N-12"x1,"N-13"x1)+boundary("N-1","N-12","N-2","N-3","N-4","N-5","N-6","N-7","N-8","N-9","N-10","","N-11","N-13");B:values("J-1"x1," j-1 "x1,"J-2"x1,"J-3"x2," J-3"x1,"j-3 "x1,"J-4"x1,"J-5"x1,"J-6"x1,"J-10"x1,"J-9"x1,""x1,"J-8"x1)+boundary("J-1",""," j-1 ","J-2","J-3"," J-3","j-3 ","J-3","J-4","J-5","J-6","J-10","J-9","J-8");C:values("Open"x9,"OPEN"x1," open "x1,"Closed"x3)+boundary("Open","Open","Open","Open","Open","OPEN","Open","Closed","Open","Closed"," open ","Open","Open","Closed")
FIXTURE  jobq_01.xlsx sheet=Hold cols=A:Job ID,B:Customer,C:Status,D:Open NCRs rows=0 stale=2
FIXTURE  jobq_01.xlsx sheet=Unmatched cols=A:NCR No,B:Job ID,C:Status rows=0 stale=2
FIXTURE  jobq_02.xlsx sheet=Jobs cols=A:Job ID,B:Customer,C:Status,D:Open NCRs rows=3 data=A:values("J-1"x1,"J-2"x1,""x1)+boundary("","J-1","J-2");B:text(6);C:values("Active"x2,""x1)+boundary("","Active","Active");D:int(8,9)+boundary("",9,9)
FIXTURE  jobq_02.xlsx sheet=Ncr cols=A:NCR No,B:Job ID,C:Status rows=4 data=A:values("N-1"x1,"N-2"x1,"N-3"x1,"N-4"x1)+boundary("N-1","N-2","N-3","N-4");B:values("J-9"x1,"J-1"x1,"J-2"x2)+boundary("J-9","J-1","J-2","J-2");C:values("Open"x3,""x1)+boundary("","Open","Open","Open")
FIXTURE  jobq_02.xlsx sheet=Hold cols=A:Job ID,B:Customer,C:Status,D:Open NCRs rows=0 stale=2
FIXTURE  jobq_02.xlsx sheet=Unmatched cols=A:NCR No,B:Job ID,C:Status rows=0 stale=2
TEST     fixture=jobq_01.xlsx call=modMain.RunHeadless(Jobs, Ncr, Hold, Unmatched) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND ROWS_USED(Hold) = 4 AND VALUE(Hold![Job ID]2) = "J-3" AND VALUE(Hold![Open NCRs]2) = 3 AND VALUE(Hold![Job ID]3) = "J-1" AND VALUE(Hold![Open NCRs]3) = 2 AND VALUE(Hold![Job ID]4) = "J-6" AND VALUE(Hold![Open NCRs]4) = 1 AND COUNTIF(Hold![Job ID], "stale") = 0 AND ROWS_USED(Unmatched) = 4 AND COUNTIF(Unmatched![Job ID], "stale") = 0
TEST     fixture=jobq_01.xlsx call=modMain.RunHeadless(Jobs, Ncr, Hold, Unmatched) assert=RESULT = 3 AND ERROR_COUNT = 0 AND VALUE(Jobs![Open NCRs]2) = 2 AND VALUE(Jobs![Open NCRs]3) = 1 AND VALUE(Jobs![Open NCRs]4) = 3 AND VALUE(Jobs![Open NCRs]5) = 9 AND VALUE(Jobs![Open NCRs]6) = 1 AND VALUE(Jobs![Open NCRs]7) = 0 AND VALUE(Jobs![Open NCRs]8) = 1 AND VALUE(Jobs![Open NCRs]9) = 0 AND COUNTA(Jobs![Job ID]) = PRE_COUNTA(Jobs![Job ID]) AND COUNTA(Ncr![NCR No]) = PRE_COUNTA(Ncr![NCR No]) AND COUNTIF(Ncr![Status], "Closed") = PRE_COUNTIF(Ncr![Status], "Closed")
TEST     fixture=jobq_01.xlsx call=modMain.RunHeadless(Jobs, Ncr, Hold, Unmatched) assert=COUNTA(Unmatched![NCR No]) = 4 AND VALUE(Unmatched![NCR No]2) = "N-12" AND VALUE(Unmatched![Status]2) = "Open" AND VALUE(Unmatched![NCR No]3) = "N-11" AND VALUE(Unmatched![Job ID]3) = "J-9" AND VALUE(Unmatched![NCR No]4) = "N-13" AND VALUE(Unmatched![Status]4) = "Closed" AND COUNTA(Unmatched![Job ID]) = 3 AND COUNTA(Unmatched![Status]) = 4 AND COUNTA(Hold![Customer]) = 4
TEST     fixture=jobq_01.xlsx call=modMain.RunHeadless(Ncr, Jobs, Hold, Unmatched) assert=RESULT = -1 AND ERROR_COUNT = 1
TEST     fixture=jobq_02.xlsx call=modMain.RunHeadless(Jobs, Ncr, Hold, Unmatched) assert=RESULT = 2 AND ERROR_COUNT = 0 AND ROWS_USED(Hold) = 3 AND VALUE(Hold![Job ID]2) = "J-2" AND VALUE(Hold![Open NCRs]2) = 2 AND VALUE(Hold![Job ID]3) = "J-1" AND VALUE(Hold![Open NCRs]3) = 1 AND ROWS_USED(Unmatched) = 2 AND VALUE(Unmatched![NCR No]2) = "N-1" AND VALUE(Unmatched![Job ID]2) = "J-9" AND COUNTA(Unmatched![Status]) = 1 AND COUNTA(Jobs![Open NCRs]) = 3
---
VCL 0.6
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
VCL 0.6
TASK
GOAL     Run all three steps of the job quality tool on the four sheets it is given and return how many jobs are on hold.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     jobs As Worksheet, ncr As Worksheet, hold As Worksheet, unmatched As Worksheet
RET      Long
VAR      n As Long
PRE      COL($jobs, "Job ID") > 0; COL($jobs, "Status") > 0; COL($jobs, "Open NCRs") > 0; COL($ncr, "NCR No") > 0; COL($ncr, "Job ID") > 0; COL($ncr, "Status") > 0
MODIFIES $jobs["Open NCRs"]
MODIFIES $hold
MODIFIES $unmatched
OP       CALL(modSteps.CountOpenNcrs, $jobs, $ncr)
OP       CALL(modSteps.WriteOnHold, $jobs, $hold, into=n)
OP       CALL(modSteps.WriteUnmatched, $jobs, $ncr, $unmatched)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write into each job's Open NCRs column how many Open NCRs the NCR register has for it, 0 when none.
PROC     CountOpenNcrs
KIND     Sub
MODULE   modSteps
ARGS     jobs As Worksheet, ncr As Worksheet
VAR      cNJob As Long
VAR      cNStatus As Long
VAR      cJob As Long
VAR      cOpen As Long
VAR      counts As Dictionary
VAR      got As Long
PRE      COL($jobs, "Job ID") > 0; COL($jobs, "Open NCRs") > 0; COL($ncr, "Job ID") > 0; COL($ncr, "Status") > 0
MODIFIES $jobs["Open NCRs"]
OP       LET(cNJob, COL($ncr, "Job ID"))
OP       LET(cNStatus, COL($ncr, "Status"))
OP       LET(cJob, COL($jobs, "Job ID"))
OP       LET(cOpen, COL($jobs, "Open NCRs"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, cNJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($ncr, r, cNJob))
OP         IF SAME(CELL($ncr, r, cNStatus), "Open")
OP           DICT_ADD(counts, TRIM(CELL($ncr, r, cNJob)), 1, mode=sum)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($jobs, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($jobs, r, cJob))
OP         IF EXISTS(counts, TRIM(CELL($jobs, r, cJob)))
OP           DICT_GET(counts, TRIM(CELL($jobs, r, cJob)), into=got)
OP         ELSE
OP           LET(got, 0)
OP         SET_VALUE(CELL($jobs, r, cOpen), got)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Jobs On Hold sheet with the jobs that are not Closed and have Open NCRs, most NCRs first, and return how many there are.
PROC     WriteOnHold
KIND     Function
MODULE   modSteps
ARGS     jobs As Worksheet, hold As Worksheet
RET      Long
VAR      cJob As Long
VAR      cStatus As Long
VAR      cOpen As Long
VAR      cLast As Long
VAR      outRow As Long
VAR      n As Long
PRE      COL($jobs, "Job ID") > 0; COL($jobs, "Status") > 0; COL($jobs, "Open NCRs") > 0
MODIFIES $hold
OP       LET(cJob, COL($jobs, "Job ID"))
OP       LET(cStatus, COL($jobs, "Status"))
OP       LET(cOpen, COL($jobs, "Open NCRs"))
OP       LET(cLast, LASTCOL($jobs, 1))
OP       CLEAR(USED_RANGE($hold), what=all)
OP       COPY_ROW($jobs, 1, $hold, 1, cols=cLast)
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($jobs, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($jobs, r, cJob))
OP         IF NOT SAME(CELL($jobs, r, cStatus), "Closed")
OP           IF IS_NUMBER(CELL($jobs, r, cOpen))
OP             IF CELL($jobs, r, cOpen) > 0
OP               COPY_ROW($jobs, r, $hold, outRow, cols=cLast)
OP               LET(outRow, outRow + 1)
OP               LET(n, n + 1)
OP       SORT(USED_RANGE($hold), key=cOpen, order=desc, header=yes)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Unmatched NCRs sheet with every NCR whose Job ID is not in the Job Tracker.
PROC     WriteUnmatched
KIND     Sub
MODULE   modSteps
ARGS     jobs As Worksheet, ncr As Worksheet, out As Worksheet
VAR      cJob As Long
VAR      cNo As Long
VAR      cNJob As Long
VAR      cLast As Long
VAR      ids As Dictionary
VAR      outRow As Long
PRE      COL($jobs, "Job ID") > 0; COL($ncr, "NCR No") > 0; COL($ncr, "Job ID") > 0
MODIFIES $out
OP       LET(cJob, COL($jobs, "Job ID"))
OP       LET(cNo, COL($ncr, "NCR No"))
OP       LET(cNJob, COL($ncr, "Job ID"))
OP       LET(cLast, LASTCOL($ncr, 1))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($jobs, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($jobs, r, cJob))
OP         DICT_ADD(ids, TRIM(CELL($jobs, r, cJob)), 1, mode=keep)
OP       CLEAR(USED_RANGE($out), what=all)
OP       COPY_ROW($ncr, 1, $out, 1, cols=cLast)
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, cNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($ncr, r, cNo))
OP         IF NOT EXISTS(ids, TRIM(CELL($ncr, r, cNJob)))
OP           COPY_ROW($ncr, r, $out, outRow, cols=cLast)
OP           LET(outRow, outRow + 1)
ERR      standard
