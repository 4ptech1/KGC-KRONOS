VCL 0.6
BUILD
# ASSUME: the real sheets are "BOM", "Job Shortage" and "Job Tracker"; the TESTs use BOM, Shortage and Tracker
# ASSUME: BOM headers (row 1) include "Job ID", "Required Qty", "Issued Qty" and "Remark"; Job Tracker headers include "Job ID", "Status" and "Material Flag"
# ASSUME: every BOM row down to the last Job ID is a line; a blank quantity is not a number, so the line gets Check
# ASSUME: a quantity typed as text that is a plain number ("9") counts as that number
# ASSUME: the Job Shortage sheet already exists; everything on it is cleared and rebuilt on every run, and a BOM line with a blank Job ID is not counted there
# ASSUME: Job IDs are matched ignoring case and spaces at the ends; jobs with the same number of Short lines come in no particular order
# ASSUME: a job is not Closed when its Status is not Closed (ignoring case and extra spaces); a blank Status counts as not Closed
# ASSUME: a Job Tracker row with a blank Job ID is not a job: its Material Flag is left as it is
# ASSUME: the message shows -1 when a heading is missing and nothing was changed
GOAL     Mark short and doubtful BOM lines, rebuild the Job Shortage sheet with the Short lines per job, and flag every open job that is waiting for material.
NAME     JobMaterial
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(bom As Worksheet, out As Worksheet, tracker As Worksheet) As Long
SIG      modSteps.MarkBomLines(bom As Worksheet)
SIG      modSteps.WriteJobShortage(bom As Worksheet, out As Worksheet)
SIG      modSteps.FlagWaitingJobs(bom As Worksheet, tracker As Worksheet) As Long
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.MarkBomLines
FLOW     modMain.RunHeadless -> modSteps.WriteJobShortage
FLOW     modMain.RunHeadless -> modSteps.FlagWaitingJobs
FIXTURE  mat_01.xlsx sheet=BOM cols=A:Job ID,B:Part No,C:Required Qty,D:Issued Qty,E:Remark rows=17 data=A:values("J-101"x4,"J-102"x2,"J-103"x4," j-103 "x1,""x1,"J-105"x1,"J-106"x2,"J-106 "x1,"J-104 "x1)+boundary("J-101","J-101","J-102","J-102","J-103","","J-103"," j-103 ","J-103","J-103","J-101","J-105","J-106","J-106","J-106 ","J-101","J-104 ");B:seq(1,1);C:int(1,20)+boundary(10,10,10,"",8,10,6,6,6,9,2,5,9,9,9,3,4);D:int(0,20)+boundary(9,10,11,5,"",2,1,0,2,4,1,5,1,2,3,8,3);E:values(""x10,"Short"x3,"old"x2,"Check"x2)+boundary("","Short","old","","Short","","","Check","","old","","Short","","","","Check","")
FIXTURE  mat_01.xlsx sheet=Shortage cols=A:Job ID,B:Short Lines rows=0 stale=6
FIXTURE  mat_01.xlsx sheet=Tracker cols=A:Job ID,B:Customer,C:Status,D:Material Flag rows=9 data=A:values("J-101"x1,"J-103"x1,"J-106"x1,"J-105"x1,""x1," j-104 "x1,"J-102"x1,"J-108"x1,"J-109"x1)+boundary("J-101","J-103","J-106","J-105",""," j-104 ","J-102","J-108","J-109");B:text(8);C:values("Open"x5,"In Progress"x1," closed "x1,""x1,"On Hold"x1)+boundary("Open","In Progress"," closed ","Open","Open","","Open","Open","On Hold");D:values(""x3,"old"x3,"Waiting material"x2,"keep"x1)+boundary("","old","Waiting material","Waiting material","keep","","old","","old")
FIXTURE  mat_02.xlsx sheet=BOM cols=A:Job ID,B:Part No,C:Required Qty,D:Issued Qty,E:Remark rows=7 data=A:values("J-201"x2,""x1,"J-202"x2,"J-203"x2)+boundary("J-201","J-201","","J-202","J-202","J-203","J-203");B:seq(1,1);C:values("10"x1,"n/a"x1,"7"x1," 12 "x1,"1e3"x1,"20"x1,"8"x1)+boundary("10","n/a","7"," 12 ","1e3","20","8");D:values("9"x1,"5"x1,"TBC"x1,"12"x1,"3"x1,"15"x1,"2"x1)+boundary("9","5","TBC","12","3","15","2");E:values(""x5,"Short"x1,"old"x1)+boundary("","Short","","old","","","")
FIXTURE  mat_02.xlsx sheet=Shortage cols=A:Job ID,B:Short Lines rows=0 stale=4
FIXTURE  mat_02.xlsx sheet=Tracker cols=A:Job ID,B:Customer,C:Status,D:Material Flag rows=4 data=A:values("J-201"x1,"J-202"x1,""x1,"J-203"x1)+boundary("J-201","J-202","","J-203");B:text(8);C:values("Open"x3,"Closed"x1)+boundary("Open","Open","Open","Closed");D:values(""x2,"old"x1,"keep"x1)+boundary("","old","keep","")
TEST     fixture=mat_01.xlsx call=modMain.RunHeadless(BOM, Shortage, Tracker) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(BOM![Remark], "Short") = 11 AND COUNTIF(BOM![Remark], "Check") = 2 AND COUNTIF(BOM![Remark], "old") = 0 AND VALUE(BOM![Remark]2) = "Short" AND VALUE(BOM![Remark]5) = "Check" AND VALUE(BOM![Remark]6) = "Check" AND VALUE(BOM![Remark]18) = "Short" AND ROWS_USED(Shortage) = 5 AND VALUE(Shortage!A1) = "Job ID" AND VALUE(Shortage!B1) = "Short Lines" AND VALUE(Shortage![Job ID]2) = "J-103" AND VALUE(Shortage![Short Lines]2) = 4 AND VALUE(Shortage![Job ID]3) = "J-106" AND VALUE(Shortage![Short Lines]3) = 3 AND VALUE(Shortage![Job ID]4) = "J-101" AND VALUE(Shortage![Short Lines]4) = 2 AND VALUE(Shortage![Job ID]5) = "J-104" AND VALUE(Shortage![Short Lines]5) = 1 AND COUNTIF(Shortage![Job ID], "stale") = 0 AND COUNTIF(Tracker![Material Flag], "Waiting material") = 3 AND VALUE(Tracker![Material Flag]2) = "Waiting material" AND VALUE(Tracker![Material Flag]3) = "Waiting material" AND VALUE(Tracker![Material Flag]7) = "Waiting material" AND COUNTIF(Tracker![Material Flag], "old") = 0 AND COUNTIF(Tracker![Material Flag], "keep") = 1
TEST     fixture=mat_01.xlsx call=modMain.RunHeadless(BOM, Shortage, Tracker) assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTA(BOM![Job ID]) = PRE_COUNTA(BOM![Job ID]) AND SUM(BOM![Required Qty]) = PRE_SUM(BOM![Required Qty]) AND SUM(BOM![Issued Qty]) = PRE_SUM(BOM![Issued Qty]) AND COUNTA(Tracker![Job ID]) = PRE_COUNTA(Tracker![Job ID]) AND COUNTA(Tracker![Status]) = PRE_COUNTA(Tracker![Status])
TEST     fixture=mat_02.xlsx call=modMain.RunHeadless(BOM, Shortage, Tracker) runs=2 assert=RESULT = 1 AND ERROR_COUNT = 0 AND VALUE(BOM![Remark]2) = "Short" AND COUNTIF(BOM![Remark], "Short") = 3 AND COUNTIF(BOM![Remark], "Check") = 3 AND COUNTIF(BOM![Remark], "old") = 0 AND ROWS_USED(Shortage) = 3 AND VALUE(Shortage![Job ID]2) = "J-203" AND VALUE(Shortage![Short Lines]2) = 2 AND VALUE(Shortage![Job ID]3) = "J-201" AND VALUE(Shortage![Short Lines]3) = 1 AND COUNTIF(Shortage![Job ID], "stale") = 0 AND VALUE(Tracker![Material Flag]2) = "Waiting material" AND COUNTIF(Tracker![Material Flag], "Waiting material") = 1 AND VALUE(Tracker![Material Flag]4) = "keep" AND COUNTIF(Tracker![Material Flag], "old") = 0 AND COUNTA(BOM![Required Qty]) = PRE_COUNTA(BOM![Required Qty])
TEST     fixture=mat_01.xlsx call=modMain.RunHeadless(Tracker, Shortage, BOM) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
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
VCL 0.6
TASK
GOAL     Run the three steps of the job material tool on the sheets it is given and return how many jobs are waiting for material.
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
OP       CALL(modSteps.MarkBomLines, $bom)
OP       CALL(modSteps.WriteJobShortage, $bom, $out)
OP       CALL(modSteps.FlagWaitingJobs, $bom, $tracker, into=waiting)
OP       RETURN(waiting)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Short in Remark on each BOM line issued less than required, Check where a quantity is not a number, and clear Remark on the other lines.
PROC     MarkBomLines
KIND     Sub
MODULE   modSteps
ARGS     bom As Worksheet
VAR      cJob As Long
VAR      cReq As Long
VAR      cIss As Long
VAR      cRemark As Long
VAR      mark As String
PRE      COL($bom, "Job ID") > 0; COL($bom, "Required Qty") > 0; COL($bom, "Issued Qty") > 0; COL($bom, "Remark") > 0
MODIFIES $bom["Remark"]
OP       LET(cJob, COL($bom, "Job ID"))
OP       LET(cReq, COL($bom, "Required Qty"))
OP       LET(cIss, COL($bom, "Issued Qty"))
OP       LET(cRemark, COL($bom, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($bom, cJob) var=r
OP         LET(mark, "Check")
OP         IF IS_NUMBER(CELL($bom, r, cReq))
OP           IF IS_NUMBER(CELL($bom, r, cIss))
OP             LET(mark, "")
OP             IF CELL($bom, r, cIss) - CELL($bom, r, cReq) < 0 THEN LET(mark, "Short")
OP         SET_VALUE(CELL($bom, r, cRemark), mark)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Job Shortage sheet with one row for each Job ID that has Short BOM lines and its number of Short lines, the most first.
PROC     WriteJobShortage
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
OP         IF SAME(CELL($bom, r, cRemark), "Short") THEN DICT_ADD(counts, TRIM(CELL($bom, r, cJob)), 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=job in=KEYS(counts)
OP         DICT_GET(counts, job, into=got)
OP         WRITE_ROW($out, outRow, job, got)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=2, order=desc, header=yes)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Waiting material in the Material Flag of every job that is not Closed and has a Short BOM line, clear the other flags and return how many jobs are waiting.
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
VAR      mark As String
VAR      n As Long
PRE      COL($bom, "Job ID") > 0; COL($bom, "Remark") > 0; COL($tracker, "Job ID") > 0; COL($tracker, "Status") > 0; COL($tracker, "Material Flag") > 0
MODIFIES $tracker["Material Flag"]
OP       LET(cJob, COL($bom, "Job ID"))
OP       LET(cRemark, COL($bom, "Remark"))
OP       LET(tJob, COL($tracker, "Job ID"))
OP       LET(tStatus, COL($tracker, "Status"))
OP       LET(tFlag, COL($tracker, "Material Flag"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($bom, cJob) var=r
OP         IF SAME(CELL($bom, r, cRemark), "Short") THEN DICT_ADD(shortJobs, TRIM(CELL($bom, r, cJob)), 1, mode=overwrite)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($tracker, tJob) var=t
OP         SKIP_IF IS_EMPTY(CELL($tracker, t, tJob))
OP         LET(mark, "")
OP         IF NOT SAME(CELL($tracker, t, tStatus), "Closed") AND EXISTS(shortJobs, TRIM(CELL($tracker, t, tJob))) THEN LET(mark, "Waiting material")
OP         SET_VALUE(CELL($tracker, t, tFlag), mark)
OP         IF mark <> "" THEN LET(n, n + 1)
OP       RETURN(n)
ERR      standard
