VCL 0.6
BUILD
# ASSUME: the real sheets are Dispatch, Job Tracker, Late Dispatches and Customer Late; the TESTs use Dispatch, JobTracker, LateDisp and CustLate
# ASSUME: Late Dispatches and Customer Late already exist and are rebuilt on every run
# ASSUME: Late needs two real dates with Actual Date after Planned Date; Overdue needs an empty Actual Date and a real Planned Date before today; every other Remark is cleared
# ASSUME: Job IDs are matched ignoring case and stray spaces; a Job Tracker row with a blank Job ID is left alone
# ASSUME: a job's Late dispatch is one whose Remark is Late after step 1; a job with no dispatch is Not shipped
# ASSUME: Customer Late lists every customer of the Job Tracker (names matched ignoring case and spaces, shown as first typed) with 0 when it has no Shipped late job; equal counts come in no fixed order
GOAL     Mark late and overdue dispatches, set each job's dispatch status, list the late and overdue dispatches by planned date and count shipped-late jobs per customer.
NAME     DispatchMonthEnd
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modDispatch
SIG      modMain.Run()
SIG      modMain.RunHeadless(disp As Worksheet, jobs As Worksheet, lateOut As Worksheet, custOut As Worksheet) As Long
SIG      modDispatch.MarkRemarks(disp As Worksheet) As Long
SIG      modDispatch.SetJobStatus(jobs As Worksheet, disp As Worksheet)
SIG      modDispatch.ListLateDispatches(disp As Worksheet, out As Worksheet)
SIG      modDispatch.CountCustomerLate(jobs As Worksheet, out As Worksheet)
FLOW     modMain.Run -> modMain.RunHeadless -> modDispatch.MarkRemarks
FLOW     modMain.RunHeadless -> modDispatch.SetJobStatus
FLOW     modMain.RunHeadless -> modDispatch.ListLateDispatches
FLOW     modMain.RunHeadless -> modDispatch.CountCustomerLate
FIXTURE  dispatch_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Job ID,C:Planned Date,D:Actual Date,E:Remark rows=10 data=A:seq(1,1);B:values("J-1"x1," j-1"x1,"J-2"x2,"J-3"x2,"J-4 "x1,"J-5"x1,""x1,"J-7"x1)+boundary("J-1"," j-1","J-2","J-2","J-3","J-3","J-4 ","J-5","","J-7");C:reldates(-30,10)+boundary(-10,-5,-4,-3,0,5,-1,-7,1,-20);D:reldates(-30,10)+boundary(-8,-5,-6,"","","","",-1,2,-19);E:values("old"x8,"Late"x1,"Overdue"x1)+boundary("old","old","old","Late","old","old","old","old","old","Overdue")
FIXTURE  dispatch_01.xlsx sheet=JobTracker cols=A:Job ID,B:Customer,C:Status,D:Dispatch Status rows=9 data=A:values(" j-1"x1,"J-2"x1,""x1,"J-3"x1,"J-4 "x1,"j-5"x1,"J-6"x1,"J-7"x1,"J-8"x1)+boundary(" j-1","J-2","","J-3","J-4 ","j-5","J-6","J-7","J-8");B:values("Acme"x2,"acme "x1,"Dyer"x1,"Bolt"x2,""x1,"Cask"x2)+boundary("Acme","acme ","Dyer","Bolt","","Cask","Cask","Acme","Bolt");C:values("Open"x9);D:values("old"x9)
FIXTURE  dispatch_01.xlsx sheet=LateDisp cols=A:Dispatch No,B:Job ID,C:Planned Date,D:Actual Date,E:Remark rows=0 stale=3
FIXTURE  dispatch_01.xlsx sheet=CustLate cols=A:Customer,B:Late Jobs rows=0 stale=2
TEST     fixture=dispatch_01.xlsx call=modMain.RunHeadless(Dispatch, JobTracker, LateDisp, CustLate) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND COUNTIF(Dispatch![Remark], "Late") = 4 AND COUNTIF(Dispatch![Remark], "Overdue") = 2 AND COUNTIF(Dispatch![Remark], "old") = 0 AND COUNTIF(JobTracker![Dispatch Status], "Shipped late") = 3 AND COUNTIF(JobTracker![Dispatch Status], "Shipped") = 1 AND COUNTIF(JobTracker![Dispatch Status], "Not shipped") = 4 AND COUNTIF(JobTracker![Dispatch Status], "old") = 1 AND COUNTA(LateDisp![Dispatch No]) = 7 AND ROWS_USED(LateDisp) = 7 AND ROWS_USED(CustLate) = 5 AND COUNTIF(LateDisp![Dispatch No], "stale") = 0 AND COUNTA(CustLate![Customer]) = 5 AND COUNTIF(CustLate![Customer], "stale") = 0
TEST     fixture=dispatch_01.xlsx call=modMain.RunHeadless(Dispatch, JobTracker, LateDisp, CustLate) assert=VALUE(LateDisp![Dispatch No]2) = 10 AND VALUE(LateDisp![Dispatch No]3) = 1 AND VALUE(LateDisp![Dispatch No]4) = 8 AND VALUE(LateDisp![Dispatch No]5) = 4 AND VALUE(LateDisp![Dispatch No]6) = 7 AND VALUE(LateDisp![Dispatch No]7) = 9 AND VALUE(LateDisp![Remark]2) = "Late" AND VALUE(LateDisp![Remark]5) = "Overdue" AND VALUE(CustLate![Customer]2) = "Acme" AND VALUE(CustLate![Late Jobs]2) = 2 AND VALUE(CustLate![Customer]3) = "Cask" AND VALUE(CustLate![Late Jobs]3) = 1 AND ERROR_COUNT = 0 AND COUNTA(Dispatch![Dispatch No]) = PRE_COUNTA(Dispatch![Dispatch No]) AND COUNTA(JobTracker![Job ID]) = PRE_COUNTA(JobTracker![Job ID])
TEST     fixture=dispatch_01.xlsx call=modMain.RunHeadless(JobTracker, Dispatch, LateDisp, CustLate) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
TASK
GOAL     Start the dispatch month-end report on the workbook's own sheets and tell the user how many dispatches are overdue.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "Dispatch"), SHEET(THIS_BOOK, "Job Tracker"), SHEET(THIS_BOOK, "Late Dispatches"), SHEET(THIS_BOOK, "Customer Late"), into=n)
OP       MSG("Overdue dispatches: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run all four report steps on the four sheets it is given and return how many dispatches are overdue.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     disp As Worksheet, jobs As Worksheet, lateOut As Worksheet, custOut As Worksheet
RET      Long
VAR      overdue As Long
PRE      COL($disp, "Dispatch No") > 0; COL($disp, "Job ID") > 0; COL($disp, "Planned Date") > 0; COL($disp, "Actual Date") > 0; COL($disp, "Remark") > 0; COL($jobs, "Job ID") > 0; COL($jobs, "Customer") > 0; COL($jobs, "Dispatch Status") > 0
MODIFIES $disp["Remark"]
MODIFIES $jobs["Dispatch Status"]
MODIFIES $lateOut
MODIFIES $custOut
OP       CALL(modDispatch.MarkRemarks, $disp, into=overdue)
OP       CALL(modDispatch.SetJobStatus, $jobs, $disp)
OP       CALL(modDispatch.ListLateDispatches, $disp, $lateOut)
OP       CALL(modDispatch.CountCustomerLate, $jobs, $custOut)
OP       RETURN(overdue)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Late, Overdue or nothing in each dispatch's Remark and return how many dispatches are overdue.
PROC     MarkRemarks
KIND     Function
MODULE   modDispatch
ARGS     disp As Worksheet
RET      Long
VAR      cNo As Long
VAR      cPlan As Long
VAR      cAct As Long
VAR      cRem As Long
VAR      n As Long
PRE      COL($disp, "Dispatch No") > 0; COL($disp, "Planned Date") > 0; COL($disp, "Actual Date") > 0; COL($disp, "Remark") > 0
MODIFIES $disp["Remark"]
OP       LET(cNo, COL($disp, "Dispatch No"))
OP       LET(cPlan, COL($disp, "Planned Date"))
OP       LET(cAct, COL($disp, "Actual Date"))
OP       LET(cRem, COL($disp, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($disp, cNo) var=r
OP         SET_VALUE(CELL($disp, r, cRem), "")
OP         IF IS_DATE(CELL($disp, r, cAct))
OP           IF IS_DATE(CELL($disp, r, cPlan))
OP             IF CELL($disp, r, cAct) > CELL($disp, r, cPlan)
OP               SET_VALUE(CELL($disp, r, cRem), "Late")
OP         ELSE
OP           IF IS_EMPTY(CELL($disp, r, cAct))
OP             IF IS_DATE(CELL($disp, r, cPlan))
OP               IF CELL($disp, r, cPlan) < TODAY
OP                 SET_VALUE(CELL($disp, r, cRem), "Overdue")
OP                 LET(n, n + 1)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Set each job's Dispatch Status to Shipped late, Shipped or Not shipped from the dispatches that carry its Job ID.
PROC     SetJobStatus
KIND     Sub
MODULE   modDispatch
ARGS     jobs As Worksheet, disp As Worksheet
VAR      cJob As Long
VAR      cStat As Long
VAR      dJob As Long
VAR      dAct As Long
VAR      dRem As Long
VAR      lateJobs As Dictionary
VAR      shipJobs As Dictionary
PRE      COL($jobs, "Job ID") > 0; COL($jobs, "Dispatch Status") > 0; COL($disp, "Job ID") > 0; COL($disp, "Actual Date") > 0; COL($disp, "Remark") > 0
MODIFIES $jobs["Dispatch Status"]
OP       LET(cJob, COL($jobs, "Job ID"))
OP       LET(cStat, COL($jobs, "Dispatch Status"))
OP       LET(dJob, COL($disp, "Job ID"))
OP       LET(dAct, COL($disp, "Actual Date"))
OP       LET(dRem, COL($disp, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($disp, dJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($disp, r, dJob))
OP         IF SAME(CELL($disp, r, dRem), "Late") THEN DICT_ADD(lateJobs, CLEAN_TEXT(CELL($disp, r, dJob)), 1, mode=keep)
OP         IF IS_DATE(CELL($disp, r, dAct)) THEN DICT_ADD(shipJobs, CLEAN_TEXT(CELL($disp, r, dJob)), 1, mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($jobs, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($jobs, r, cJob))
OP         IF EXISTS(lateJobs, CLEAN_TEXT(CELL($jobs, r, cJob)))
OP           SET_VALUE(CELL($jobs, r, cStat), "Shipped late")
OP         ELSE
OP           IF EXISTS(shipJobs, CLEAN_TEXT(CELL($jobs, r, cJob)))
OP             SET_VALUE(CELL($jobs, r, cStat), "Shipped")
OP           ELSE
OP             SET_VALUE(CELL($jobs, r, cStat), "Not shipped")
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Late Dispatches sheet with the Late and Overdue dispatches under the same headings, earliest planned date first.
PROC     ListLateDispatches
KIND     Sub
MODULE   modDispatch
ARGS     disp As Worksheet, out As Worksheet
VAR      cRem As Long
VAR      cPlan As Long
VAR      nCols As Long
VAR      outRow As Long
PRE      COL($disp, "Planned Date") > 0; COL($disp, "Remark") > 0
MODIFIES $out
OP       LET(cRem, COL($disp, "Remark"))
OP       LET(cPlan, COL($disp, "Planned Date"))
OP       LET(nCols, LASTCOL($disp, 1))
OP       CLEAR(USED_RANGE($out), what=all)
OP       COPY_ROW($disp, 1, $out, 1, cols=nCols)
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($disp, cPlan) var=r
OP         IF SAME(CELL($disp, r, cRem), "Late") OR SAME(CELL($disp, r, cRem), "Overdue")
OP           COPY_ROW($disp, r, $out, outRow, cols=nCols)
OP           LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=cPlan, order=asc, header=yes)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Customer Late sheet with each customer and its number of Shipped late jobs, the most first.
PROC     CountCustomerLate
KIND     Sub
MODULE   modDispatch
ARGS     jobs As Worksheet, out As Worksheet
VAR      cCust As Long
VAR      cStat As Long
VAR      counts As Dictionary
VAR      got As Long
VAR      outRow As Long
PRE      COL($jobs, "Customer") > 0; COL($jobs, "Dispatch Status") > 0
MODIFIES $out
OP       LET(cCust, COL($jobs, "Customer"))
OP       LET(cStat, COL($jobs, "Dispatch Status"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Customer", "Late Jobs")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($jobs, cCust) var=r
OP         SKIP_IF IS_EMPTY(CELL($jobs, r, cCust))
OP         DICT_ADD(counts, CLEAN_TEXT(CELL($jobs, r, cCust)), 0, mode=sum)
OP         IF SAME(CELL($jobs, r, cStat), "Shipped late") THEN DICT_ADD(counts, CLEAN_TEXT(CELL($jobs, r, cCust)), 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         WRITE_ROW($out, outRow, k, got)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=2, order=desc, header=yes)
ERR      standard
