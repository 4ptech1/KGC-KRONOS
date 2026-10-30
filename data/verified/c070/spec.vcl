VCL 0.6
BUILD
# ASSUME: Dispatch headers (row 1) include "Dispatch No", "Job ID", "Planned Date", "Actual Date" and "Remark"; Job Tracker headers include "Job ID" and "Dispatch Check"
# ASSUME: the Remark and Dispatch Check columns already exist; the tool fills and clears them but never adds them
# ASSUME: a dispatch is Late only when both dates are real dates and the Actual Date is after the Planned Date; every other Remark is cleared
# ASSUME: a row with a blank Dispatch No is not a dispatch, so no step marks, copies or counts it
# ASSUME: a job is Dispatched when a dispatch row with its Job ID (trimmed, case ignored) holds a real date as Actual Date
# ASSUME: Late Dispatches gets the Dispatch headings and the values of every column; it and Dispatch Summary exist and are rebuilt on every run
# ASSUME: the real sheets are "Dispatch", "Late Dispatches", "Job Tracker" and "Dispatch Summary"; the TESTs use one-word names
GOAL     Mark late dispatches, copy them to a Late Dispatches sheet, mark dispatched jobs in the Job Tracker and rebuild a dispatch summary.
NAME     DispatchReview
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(disp As Worksheet, lateOut As Worksheet, jobs As Worksheet, summ As Worksheet) As Long
SIG      modSteps.MarkLateDispatches(disp As Worksheet) As Long
SIG      modSteps.CopyLateRows(disp As Worksheet, lateOut As Worksheet)
SIG      modSteps.MarkDispatchedJobs(disp As Worksheet, jobs As Worksheet)
SIG      modSteps.WriteDispatchSummary(disp As Worksheet, jobs As Worksheet, summ As Worksheet)
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.MarkLateDispatches
FLOW     modMain.RunHeadless -> modSteps.CopyLateRows
FLOW     modMain.RunHeadless -> modSteps.MarkDispatchedJobs
FLOW     modMain.RunHeadless -> modSteps.WriteDispatchSummary
FIXTURE  dispatch_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Job ID,C:Customer,D:Planned Date,E:Actual Date,F:Remark,G:Notes rows=10 data=A:seq(5001,1)+boundary(5001,5002,5003,5004,5005,"",5007,5008,5009,5010);B:values("J-101"x1,"J-102"x2," J-103 "x1,"J-104"x2,"J-105"x1,""x1,"J-106"x1,"J-107"x1)+boundary("J-101","J-102"," J-103 ","J-104","J-104","J-105","","J-102","J-106","J-107");C:text(8);D:dates(2026-03-01,2026-03-31)+boundary(2026-03-10,2026-03-10,2026-03-12,"",2026-03-14,2026-03-01,2026-03-05,2026-03-20,2026-03-18,2026-03-22);E:dates(2026-03-01,2026-03-31)+boundary(2026-03-11,2026-03-10,2026-03-11,2026-03-15,"",2026-03-20,2026-03-09,2026-03-25,"",2026-03-24);F:values("Late"x4,"old"x2,""x4)+boundary("","Late","old","","Late","Late","","Late","old","");G:values("chk"x4,""x6)+boundary("","chk","","chk","","","chk","chk","","")
FIXTURE  dispatch_01.xlsx sheet=LateDispatches cols=A:Dispatch No,B:Job ID,C:Customer,D:Planned Date,E:Actual Date,F:Remark,G:Notes rows=0 stale=6
FIXTURE  dispatch_01.xlsx sheet=JobTracker cols=A:Job ID,B:Status,C:Dispatch Check rows=9 data=A:values("J-101"x1,"J-102"x1,"J-103"x1,"J-104 "x1,""x1,"J-105"x1,"J-106"x1,"j-107"x1,"J-108"x1)+boundary("J-101","J-102","J-103","J-104 ","","J-105","J-106","j-107","J-108");B:values("Open"x5,"Closed"x4);C:values("Dispatched"x3,"old"x2,""x4)+boundary("","old","","","Dispatched","Dispatched","old","","Dispatched")
FIXTURE  dispatch_01.xlsx sheet=DispatchSummary cols=A:Measure,B:Count rows=0 stale=5
TEST     fixture=dispatch_01.xlsx call=modMain.RunHeadless(Dispatch, LateDispatches, JobTracker, DispatchSummary) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(Dispatch![Remark], "Late") = 4 AND COUNTA(Dispatch![Remark]) = 5 AND ROWS_USED(LateDispatches) = 5 AND COUNTIF(LateDispatches![Remark], "Late") = 4 AND COUNTA(LateDispatches![Notes]) = 3 AND COUNTIF(JobTracker![Dispatch Check], "Dispatched") = 5 AND COUNTA(JobTracker![Dispatch Check]) = 6 AND ROWS_USED(DispatchSummary) = 4 AND VALUE(DispatchSummary![Count]2) = 9 AND VALUE(DispatchSummary![Count]3) = 4 AND VALUE(DispatchSummary![Count]4) = 5
TEST     fixture=dispatch_01.xlsx call=modMain.RunHeadless(Dispatch, LateDispatches, JobTracker, DispatchSummary) assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(Dispatch![Dispatch No]) = PRE_COUNTA(Dispatch![Dispatch No]) AND VALUE(Dispatch![Remark]2) = "Late" AND VALUE(Dispatch![Remark]11) = "Late" AND VALUE(LateDispatches![Dispatch No]2) = 5001 AND VALUE(LateDispatches![Dispatch No]5) = 5010 AND VALUE(JobTracker![Dispatch Check]2) = "Dispatched" AND VALUE(JobTracker![Dispatch Check]5) = "Dispatched" AND VALUE(JobTracker![Dispatch Check]9) = "Dispatched" AND VALUE(DispatchSummary![Measure]2) = "Dispatches" AND VALUE(DispatchSummary![Measure]3) = "Late dispatches" AND VALUE(DispatchSummary![Measure]4) = "Dispatched jobs" AND FILL_COUNT(LateDispatches!A1:G9, #FF0000) = 0 AND FILL_COUNT(DispatchSummary!A1:B9, #FF0000) = 0
TEST     fixture=dispatch_01.xlsx call=modMain.RunHeadless(JobTracker, LateDispatches, Dispatch, DispatchSummary) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
TASK
GOAL     Start the dispatch review on the workbook's own four sheets and tell the user how many dispatches are late.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "Dispatch"), SHEET(THIS_BOOK, "Late Dispatches"), SHEET(THIS_BOOK, "Job Tracker"), SHEET(THIS_BOOK, "Dispatch Summary"), into=n)
OP       MSG("Late dispatches: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run the four dispatch review steps on the four sheets it is given and return how many dispatches are late.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     disp As Worksheet, lateOut As Worksheet, jobs As Worksheet, summ As Worksheet
RET      Long
VAR      nLate As Long
PRE      COL($disp, "Dispatch No") > 0; COL($disp, "Job ID") > 0; COL($disp, "Planned Date") > 0; COL($disp, "Actual Date") > 0; COL($disp, "Remark") > 0; COL($jobs, "Job ID") > 0; COL($jobs, "Dispatch Check") > 0
MODIFIES $disp["Remark"]
MODIFIES $lateOut
MODIFIES $jobs["Dispatch Check"]
MODIFIES $summ
OP       CALL(modSteps.MarkLateDispatches, $disp, into=nLate)
OP       CALL(modSteps.CopyLateRows, $disp, $lateOut)
OP       CALL(modSteps.MarkDispatchedJobs, $disp, $jobs)
OP       CALL(modSteps.WriteDispatchSummary, $disp, $jobs, $summ)
OP       RETURN(nLate)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Late in the Remark column of every dispatch whose actual date is after its planned date, clear the other Remarks and return how many are late.
PROC     MarkLateDispatches
KIND     Function
MODULE   modSteps
ARGS     disp As Worksheet
RET      Long
VAR      cNo As Long
VAR      cPlan As Long
VAR      cActual As Long
VAR      cRemark As Long
VAR      n As Long
PRE      COL($disp, "Dispatch No") > 0; COL($disp, "Planned Date") > 0; COL($disp, "Actual Date") > 0; COL($disp, "Remark") > 0
MODIFIES $disp["Remark"]
OP       LET(cNo, COL($disp, "Dispatch No"))
OP       LET(cPlan, COL($disp, "Planned Date"))
OP       LET(cActual, COL($disp, "Actual Date"))
OP       LET(cRemark, COL($disp, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($disp, cNo) var=r
OP         CLEAR(CELL($disp, r, cRemark), what=contents)
OP         SKIP_IF IS_EMPTY(CELL($disp, r, cNo))
OP         IF IS_DATE(CELL($disp, r, cActual))
OP           IF IS_DATE(CELL($disp, r, cPlan))
OP             IF CELL($disp, r, cActual) > CELL($disp, r, cPlan)
OP               SET_VALUE(CELL($disp, r, cRemark), "Late")
OP               LET(n, n + 1)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the late dispatches sheet with the Dispatch sheet's headings and a copy of every row whose Remark is Late.
PROC     CopyLateRows
KIND     Sub
MODULE   modSteps
ARGS     disp As Worksheet, lateOut As Worksheet
VAR      cNo As Long
VAR      cRemark As Long
VAR      nCols As Long
VAR      outRow As Long
PRE      COL($disp, "Dispatch No") > 0; COL($disp, "Remark") > 0
MODIFIES $lateOut
OP       LET(cNo, COL($disp, "Dispatch No"))
OP       LET(cRemark, COL($disp, "Remark"))
OP       LET(nCols, LASTCOL($disp, 1))
OP       CLEAR(USED_RANGE($lateOut), what=all)
OP       COPY_ROW($disp, 1, $lateOut, 1, cols=nCols)
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($disp, cNo) var=r
OP         IF SAME(CELL($disp, r, cRemark), "Late")
OP           COPY_ROW($disp, r, $lateOut, outRow, cols=nCols)
OP           LET(outRow, outRow + 1)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Dispatched in the Job Tracker's Dispatch Check column for every job with a dispatch row that has an actual date, and clear it for the other jobs.
PROC     MarkDispatchedJobs
KIND     Sub
MODULE   modSteps
ARGS     disp As Worksheet, jobs As Worksheet
VAR      cNo As Long
VAR      cDispJob As Long
VAR      cActual As Long
VAR      cJob As Long
VAR      cCheck As Long
VAR      shipped As Dictionary
PRE      COL($disp, "Dispatch No") > 0; COL($disp, "Job ID") > 0; COL($disp, "Actual Date") > 0; COL($jobs, "Job ID") > 0; COL($jobs, "Dispatch Check") > 0
MODIFIES $jobs["Dispatch Check"]
OP       LET(cNo, COL($disp, "Dispatch No"))
OP       LET(cDispJob, COL($disp, "Job ID"))
OP       LET(cActual, COL($disp, "Actual Date"))
OP       LET(cJob, COL($jobs, "Job ID"))
OP       LET(cCheck, COL($jobs, "Dispatch Check"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($disp, cNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($disp, r, cNo))
OP         SKIP_IF IS_EMPTY(CELL($disp, r, cDispJob))
OP         IF IS_DATE(CELL($disp, r, cActual)) THEN DICT_ADD(shipped, TRIM(CELL($disp, r, cDispJob)), 1, mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($jobs, cJob) var=j
OP         CLEAR(CELL($jobs, j, cCheck), what=contents)
OP         IF EXISTS(shipped, TRIM(CELL($jobs, j, cJob))) THEN SET_VALUE(CELL($jobs, j, cCheck), "Dispatched")
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the dispatch summary sheet with labelled counts of dispatches, late dispatches and dispatched jobs.
PROC     WriteDispatchSummary
KIND     Sub
MODULE   modSteps
ARGS     disp As Worksheet, jobs As Worksheet, summ As Worksheet
VAR      cNo As Long
VAR      cRemark As Long
VAR      cCheck As Long
VAR      nDisp As Long
VAR      nLate As Long
VAR      nJobs As Long
PRE      COL($disp, "Dispatch No") > 0; COL($disp, "Remark") > 0; COL($jobs, "Dispatch Check") > 0
MODIFIES $summ
OP       LET(cNo, COL($disp, "Dispatch No"))
OP       LET(cRemark, COL($disp, "Remark"))
OP       LET(cCheck, COL($jobs, "Dispatch Check"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($disp, cNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($disp, r, cNo))
OP         LET(nDisp, nDisp + 1)
OP         IF SAME(CELL($disp, r, cRemark), "Late") THEN LET(nLate, nLate + 1)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($jobs, cCheck) var=j
OP         IF SAME(CELL($jobs, j, cCheck), "Dispatched") THEN LET(nJobs, nJobs + 1)
OP       CLEAR(USED_RANGE($summ), what=all)
OP       WRITE_ROW($summ, 1, "Measure", "Count")
OP       WRITE_ROW($summ, 2, "Dispatches", nDisp)
OP       WRITE_ROW($summ, 3, "Late dispatches", nLate)
OP       WRITE_ROW($summ, 4, "Dispatched jobs", nJobs)
ERR      standard
