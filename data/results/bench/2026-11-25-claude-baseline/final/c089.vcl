VCL 0.7
BUILD
# ASSUME: BOM headers (row 1) include "Job ID", "Required Qty", "Issued Qty" and "Remark"; Job Tracker headers include "Job ID", "Status" and "Material Flag"
# ASSUME: the real sheets are "BOM", "Job Shortage" and "Job Tracker"; the TESTs use BOM, JobShortage and JobTracker; Job Shortage already exists and is rebuilt on every run
# ASSUME: a job is not Closed when its Status is not SAME as Closed; Job IDs are matched after TRIM; ties in the shortage list keep no particular order
GOAL     Mark short BOM lines, list the jobs with short lines by count, flag open jobs waiting for material and report how many jobs are waiting.
NAME     JobMaterial
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(bom As Worksheet, out As Worksheet, tracker As Worksheet) As Long
SIG      modSteps.MarkShortLines(bom As Worksheet)
SIG      modSteps.WriteShortage(bom As Worksheet, out As Worksheet)
SIG      modSteps.FlagWaitingJobs(bom As Worksheet, tracker As Worksheet) As Long
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.MarkShortLines
FLOW     modMain.RunHeadless -> modSteps.WriteShortage
FLOW     modMain.RunHeadless -> modSteps.FlagWaitingJobs
FIXTURE  mat_01.xlsx sheet=BOM cols=A:Job ID,B:Part No,C:Required Qty,D:Issued Qty,E:Remark rows=10 data=A:values("J-1"x2,"J-2"x2,"J-3"x3,"J-4"x2,"J-5"x1)+boundary("J-1","J-1","J-2","J-2","J-3","J-3","J-4","J-4","J-5","J-3");B:text(6);C:int(1,50)+boundary(10,10,20,5,8,8,6,30,9,7);D:values("5"x1,"9"x1,"25"x1,"n/a"x1,"2"x1,"3"x1,""x1,"30"x1,"4"x1,"1"x1)+boundary("5","9","25","n/a","2","3","","30","4","1");E:values("old"x10)
FIXTURE  mat_01.xlsx sheet=JobShortage cols=A:Job ID,B:Short Lines rows=0 stale=2
FIXTURE  mat_01.xlsx sheet=JobTracker cols=A:Job ID,B:Customer,C:Status,D:Material Flag rows=6 data=A:values("J-1"x1,"J-2"x1,"J-3"x1,"J-4"x1,"J-5"x1,"J-6"x1)+boundary("J-1","J-2","J-3","J-4","J-5","J-6");B:text(8);C:values("Open"x4,"Closed"x1,"Hold"x1)+boundary("Open","Open","Closed","Open","Hold","Open");D:values("old"x6)
FIXTURE  mat_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=mat_01.xlsx call=modMain.RunHeadless(BOM, JobShortage, JobTracker) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND COUNTIF(BOM![Remark], "Short") = 6 AND COUNTIF(BOM![Remark], "Check") = 2 AND COUNTIF(BOM![Remark], "old") = 0 AND COUNTA(JobShortage![Job ID]) = 4 AND COUNTIF(JobShortage![Job ID], "stale") = 0 AND VALUE(JobShortage![Job ID]2) = "J-3" AND VALUE(JobShortage![Short Lines]2) = 3 AND COUNTIF(JobTracker![Material Flag], "Waiting material") = 2 AND COUNTIF(JobTracker![Material Flag], "old") = 0
TEST     fixture=mat_01.xlsx call=modMain.RunHeadless(BOM, JobShortage, JobTracker) assert=RESULT = 2 AND ERROR_COUNT = 0 AND COUNTA(BOM![Job ID]) = PRE_COUNTA(BOM![Job ID]) AND COUNTA(JobTracker![Job ID]) = PRE_COUNTA(JobTracker![Job ID]) AND VALUE(JobShortage![Job ID]3) = "J-1" AND VALUE(JobShortage![Short Lines]3) = 2 AND VALUE(JobShortage![Job ID]4) = "J-5" AND VALUE(JobShortage![Short Lines]4) = 1 AND VALUE(JobTracker![Material Flag]2) = "Waiting material" AND VALUE(JobTracker![Material Flag]4) = ""
TEST     fixture=mat_01.xlsx call=modMain.RunHeadless(Notes, JobShortage, JobTracker) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.7
TASK
GOAL     Start the job material tool on the workbook's own sheets and tell the user how many jobs are waiting for material.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "BOM"), SHEET(THIS_BOOK, "Job Shortage"), SHEET(THIS_BOOK, "Job Tracker"), into=n)
OP       MSG("Jobs waiting for material: " & n)
ERR      standard
---
VCL 0.7
TASK
GOAL     Run all steps of the job material tool on the three sheets it is given and return how many jobs are waiting for material.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     bom As Worksheet, out As Worksheet, tracker As Worksheet
RET      Long
VAR      waiting As Long
PRE      COL($bom, "Job ID") > 0; COL($bom, "Required Qty") > 0; COL($bom, "Issued Qty") > 0; COL($bom, "Remark") > 0; COL($tracker, "Job ID") > 0; COL($tracker, "Status") > 0; COL($tracker, "Material Flag") > 0
MODIFIES $bom["Remark"]
MODIFIES $out
MODIFIES $tracker["Material Flag"]
OP       CALL(modSteps.MarkShortLines, $bom)
OP       CALL(modSteps.WriteShortage, $bom, $out)
OP       CALL(modSteps.FlagWaitingJobs, $bom, $tracker, into=waiting)
OP       RETURN(waiting)
ERR      standard
---
VCL 0.7
TASK
GOAL     Write Short in Remark where a BOM line was issued less than required, Check where a quantity is not a number, and clear Remark on the others.
PROC     MarkShortLines
KIND     Sub
MODULE   modSteps
ARGS     bom As Worksheet
VAR      cJob As Long
VAR      cReq As Long
VAR      cIss As Long
VAR      cRemark As Long
PRE      COL($bom, "Job ID") > 0; COL($bom, "Required Qty") > 0; COL($bom, "Issued Qty") > 0; COL($bom, "Remark") > 0
MODIFIES $bom["Remark"]
OP       LET(cJob, COL($bom, "Job ID"))
OP       LET(cReq, COL($bom, "Required Qty"))
OP       LET(cIss, COL($bom, "Issued Qty"))
OP       LET(cRemark, COL($bom, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($bom, cJob) var=r
OP         IF IS_NUMBER(CELL($bom, r, cReq))
OP           IF IS_NUMBER(CELL($bom, r, cIss))
OP             IF CELL($bom, r, cIss) < CELL($bom, r, cReq)
OP               SET_VALUE(CELL($bom, r, cRemark), "Short")
OP             ELSE
OP               SET_VALUE(CELL($bom, r, cRemark), "")
OP           ELSE
OP             SET_VALUE(CELL($bom, r, cRemark), "Check")
OP         ELSE
OP           SET_VALUE(CELL($bom, r, cRemark), "Check")
ERR      standard
---
VCL 0.7
TASK
GOAL     Rebuild the Job Shortage sheet with each job that has Short BOM lines and its number of Short lines, the most first.
PROC     WriteShortage
KIND     Sub
MODULE   modSteps
ARGS     bom As Worksheet, out As Worksheet
VAR      cJob As Long
VAR      cRemark As Long
VAR      counts As Dictionary
VAR      got As Long
VAR      outRow As Long
PRE      COL($bom, "Job ID") > 0; COL($bom, "Remark") > 0
MODIFIES $out
OP       LET(cJob, COL($bom, "Job ID"))
OP       LET(cRemark, COL($bom, "Remark"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Job ID", "Short Lines")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($bom, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($bom, r, cJob))
OP         IF SAME(CELL($bom, r, cRemark), "Short")
OP           DICT_ADD(counts, TRIM(CELL($bom, r, cJob)), 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         WRITE_ROW($out, outRow, k, got)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=2, order=desc, header=yes)
ERR      standard
---
VCL 0.7
TASK
GOAL     Write Waiting material in Material Flag for every job that is not Closed and has a Short BOM line, clear it on the others and return how many are waiting.
PROC     FlagWaitingJobs
KIND     Function
MODULE   modSteps
ARGS     bom As Worksheet, tracker As Worksheet
RET      Long
VAR      cJob As Long
VAR      cRemark As Long
VAR      tJob As Long
VAR      tStatus As Long
VAR      tFlag As Long
VAR      shortJobs As Dictionary
VAR      waiting As Long
PRE      COL($bom, "Job ID") > 0; COL($bom, "Remark") > 0; COL($tracker, "Job ID") > 0; COL($tracker, "Status") > 0; COL($tracker, "Material Flag") > 0
MODIFIES $tracker["Material Flag"]
OP       LET(cJob, COL($bom, "Job ID"))
OP       LET(cRemark, COL($bom, "Remark"))
OP       LET(tJob, COL($tracker, "Job ID"))
OP       LET(tStatus, COL($tracker, "Status"))
OP       LET(tFlag, COL($tracker, "Material Flag"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($bom, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($bom, r, cJob))
OP         IF SAME(CELL($bom, r, cRemark), "Short")
OP           DICT_ADD(shortJobs, TRIM(CELL($bom, r, cJob)), 1, mode=sum)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($tracker, tJob) var=t
OP         IF NOT SAME(CELL($tracker, t, tStatus), "Closed") AND EXISTS(shortJobs, TRIM(CELL($tracker, t, tJob)))
OP           SET_VALUE(CELL($tracker, t, tFlag), "Waiting material")
OP           LET(waiting, waiting + 1)
OP         ELSE
OP           SET_VALUE(CELL($tracker, t, tFlag), "")
OP       RETURN(waiting)
ERR      standard
