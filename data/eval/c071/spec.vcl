VCL 0.6
BUILD
# ASSUME: the real sheets are "NDE Report", "Weld Log", "Repairs" and "Method Summary"; the TESTs use NDE, WeldLog, Repairs and MethodSummary
# ASSUME: NDE Report headers (row 1) include "Weld Seam No", "Method" and "Result"; Weld Log headers include "Weld Seam No"
# ASSUME: a missing "Action" (NDE Report) or "NDE Result" (Weld Log) header is added in row 1 just right of the last header, over an empty column
# ASSUME: Repairs and Method Summary already exist; both are cleared and rebuilt on every run
# ASSUME: the NDE rows run from row 2 to the last Weld Seam No; Fail is matched ignoring case and extra spaces
# ASSUME: seams are matched after TRIM, ignoring case; the NDE Report is in date order, so a seam's last row with a Result gives its NDE Result
# ASSUME: a row with a blank Weld Seam No still counts in steps 1, 2 and 4; a row with a blank Method is left out of the Method Summary
# ASSUME: each sheet holds well under 1,000 rows
GOAL     Mark failed NDE reports for repair, list the repairs, copy each seam's NDE result into the Weld Log and rebuild a summary of reports and fails by NDE method.
NAME     NdeFollowUp
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(nde As Worksheet, weldLog As Worksheet, repairs As Worksheet, summary As Worksheet) As Long
SIG      modSteps.MarkRepairs(nde As Worksheet) As Long
SIG      modSteps.ListRepairs(nde As Worksheet, repairs As Worksheet)
SIG      modSteps.FillNdeResults(nde As Worksheet, weldLog As Worksheet)
SIG      modSteps.WriteMethodSummary(nde As Worksheet, summary As Worksheet)
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.MarkRepairs
FLOW     modMain.RunHeadless -> modSteps.ListRepairs
FLOW     modMain.RunHeadless -> modSteps.FillNdeResults
FLOW     modMain.RunHeadless -> modSteps.WriteMethodSummary
FIXTURE  nde_01.xlsx sheet=NDE cols=A:Weld Seam No,B:Method,C:Result,D:Report Date,E:Action rows=10 data=A:values("WS-01"x1,"WS-02"x3," WS-03 "x1,""x1,"WS-04"x1,"WS-05"x1,"WS-07"x1,"WS-06"x1)+boundary("WS-01","WS-02"," WS-03 ","","WS-02","WS-04","WS-02","WS-05","WS-07","WS-06");B:values("RT"x3,"UT"x2,"rt"x1,"MT"x3,""x1)+boundary("RT","UT","rt","MT","RT","","UT","MT","MT","RT");C:values("Fail"x5," fail "x1,"Pass"x3,""x1)+boundary("Fail","Fail"," fail ","Fail","Pass","Fail","","Pass","Pass","Fail");D:dates(2026-01-05,2026-09-30);E:values(""x5,"old"x2,"Repair"x3)+boundary("","old","","","Repair","","old","Repair","Repair","")
FIXTURE  nde_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Welder ID,C:NDE Result rows=7 data=A:values("WS-01"x1,"ws-03"x1,"WS-08"x1,""x1,"WS-02"x1," WS-06"x1,"WS-04"x1)+boundary("WS-01","ws-03","WS-08","","WS-02"," WS-06","WS-04");B:values("W-11"x2,"W-12"x2,"W-13"x1,"W-14"x1,"W-15"x1);C:values("old"x3,""x2,"Pass"x1,"Fail"x1)+boundary("old","","Pass","old","Fail","","old")
FIXTURE  nde_01.xlsx sheet=Repairs cols=A:Weld Seam No,B:Method rows=0 stale=9
FIXTURE  nde_01.xlsx sheet=MethodSummary cols=A:Method,B:Reports,C:Fails rows=0 stale=5
FIXTURE  nde_02.xlsx sheet=NDE cols=A:Weld Seam No,B:Method,C:Result,D:Report Date rows=5 data=A:values("WS-21"x2,""x1,"WS-22"x1,"WS-23"x1)+boundary("WS-21","","WS-22","WS-23","WS-21");B:values("RT"x2,"UT"x1,""x1,"MT"x1)+boundary("RT","UT","","MT","RT");C:values("Fail"x2,"Pass"x2,""x1)+boundary("Fail","Pass","Fail","","Pass");D:dates(2026-02-02,2026-08-28)+boundary("")
FIXTURE  nde_02.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Welder ID rows=3 data=A:values("WS-22"x1,"WS-23"x1,"WS-21"x1)+boundary("WS-22","WS-23","WS-21");B:values(""x1,"W-22"x1,"W-23"x1)+boundary("","W-22","W-23")
FIXTURE  nde_02.xlsx sheet=Repairs cols=A:Weld Seam No,B:Method rows=0 stale=1
FIXTURE  nde_02.xlsx sheet=MethodSummary cols=A:Method,B:Reports,C:Fails rows=0 stale=1
TEST     fixture=nde_01.xlsx call=modMain.RunHeadless(NDE, WeldLog, Repairs, MethodSummary) runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND COUNTIF(NDE![Action], "Repair") = 6 AND COUNTA(NDE![Action]) = 7 AND ROWS_USED(Repairs) = 7 AND COUNTIF(Repairs![Weld Seam No], "stale") = 0 AND FILL_COUNT(Repairs!A1:B12, #FF0000) = 0 AND COUNTIF(WeldLog![NDE Result], "Fail") = 4 AND COUNTIF(WeldLog![NDE Result], "Pass") = 1 AND COUNTA(WeldLog![NDE Result]) = 6 AND ROWS_USED(MethodSummary) = 4 AND COUNTIF(MethodSummary![Method], "stale") = 0 AND FILL_COUNT(MethodSummary!A1:C8, #FF0000) = 0 AND SUM(MethodSummary![Reports]) = 9 AND SUM(MethodSummary![Fails]) = 5
TEST     fixture=nde_01.xlsx call=modMain.RunHeadless(NDE, WeldLog, Repairs, MethodSummary) assert=RESULT = 6 AND ERROR_COUNT = 0 AND COUNTA(NDE![Weld Seam No]) = PRE_COUNTA(NDE![Weld Seam No]) AND COUNTA(NDE![Method]) = PRE_COUNTA(NDE![Method]) AND COUNTIF(NDE![Result], "Fail") = PRE_COUNTIF(NDE![Result], "Fail") AND COUNTA(NDE![Report Date]) = PRE_COUNTA(NDE![Report Date]) AND COUNTA(WeldLog![Weld Seam No]) = PRE_COUNTA(WeldLog![Weld Seam No]) AND COUNTA(WeldLog![Welder ID]) = PRE_COUNTA(WeldLog![Welder ID]) AND COUNTIF(NDE![Action], "old") = 0 AND COUNTIF(WeldLog![NDE Result], "old") = 0 AND VALUE(NDE![Action]4) = "Repair" AND VALUE(NDE![Action]11) = "Repair"
TEST     fixture=nde_01.xlsx call=modMain.RunHeadless(NDE, WeldLog, Repairs, MethodSummary) assert=VALUE(Repairs![Weld Seam No]2) = "WS-01" AND VALUE(Repairs![Method]3) = "UT" AND VALUE(Repairs![Weld Seam No]7) = "WS-06" AND COUNTA(Repairs![Method]) = 6 AND VALUE(WeldLog![NDE Result]2) = "Fail" AND VALUE(WeldLog![NDE Result]3) = "fail" AND VALUE(WeldLog![NDE Result]6) = "Pass" AND VALUE(WeldLog![NDE Result]8) = "Fail" AND VALUE(MethodSummary![Method]2) = "RT" AND VALUE(MethodSummary![Reports]2) = 4 AND VALUE(MethodSummary![Fails]2) = 3 AND VALUE(MethodSummary![Method]3) = "UT" AND VALUE(MethodSummary![Fails]3) = 1 AND VALUE(MethodSummary![Method]4) = "MT" AND VALUE(MethodSummary![Reports]4) = 3
TEST     fixture=nde_01.xlsx call=modMain.RunHeadless(WeldLog, NDE, Repairs, MethodSummary) assert=RESULT = -1 AND ERROR_COUNT = 1
TEST     fixture=nde_02.xlsx call=modMain.RunHeadless(NDE, WeldLog, Repairs, MethodSummary) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(NDE!E1) = "Action" AND COUNTA(NDE!A1:J1) = 5 AND COUNTIF(NDE!E2:E9, "Repair") = 2 AND VALUE(NDE!E4) = "Repair" AND VALUE(WeldLog!C1) = "NDE Result" AND COUNTA(WeldLog!A1:H1) = 3 AND COUNTIF(WeldLog!C2:C9, "Fail") = 1 AND VALUE(WeldLog!C4) = "Pass" AND COUNTA(WeldLog!C2:C9) = 2 AND ROWS_USED(Repairs) = 3 AND SUM(MethodSummary![Reports]) = 4 AND SUM(MethodSummary![Fails]) = 1
---
VCL 0.6
TASK
GOAL     Start the NDE follow-up on the workbook's own sheets and tell the user how many failed NDE reports were marked for repair.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "NDE Report"), SHEET(THIS_BOOK, "Weld Log"), SHEET(THIS_BOOK, "Repairs"), SHEET(THIS_BOOK, "Method Summary"), into=n)
OP       MSG("NDE reports marked for repair: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run all four NDE follow-up steps on the sheets it is given and return how many failed NDE reports were marked for repair.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     nde As Worksheet, weldLog As Worksheet, repairs As Worksheet, summary As Worksheet
RET      Long
VAR      marked As Long
PRE      COL($nde, "Weld Seam No") > 0; COL($nde, "Method") > 0; COL($nde, "Result") > 0; COL($weldLog, "Weld Seam No") > 0
MODIFIES $nde
MODIFIES $weldLog
MODIFIES $repairs
MODIFIES $summary
OP       CALL(modSteps.MarkRepairs, $nde, into=marked)
OP       CALL(modSteps.ListRepairs, $nde, $repairs)
OP       CALL(modSteps.FillNdeResults, $nde, $weldLog)
OP       CALL(modSteps.WriteMethodSummary, $nde, $summary)
OP       RETURN(marked)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Repair in the Action column of every failed NDE report and clear it on every other report, adding the Action header when it is missing.
PROC     MarkRepairs
KIND     Function
MODULE   modSteps
ARGS     nde As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cResult As Long
VAR      cAction As Long
VAR      n As Long
PRE      COL($nde, "Weld Seam No") > 0; COL($nde, "Result") > 0
MODIFIES $nde
OP       LET(cSeam, COL($nde, "Weld Seam No"))
OP       LET(cResult, COL($nde, "Result"))
OP       LET(cAction, COL($nde, "Action"))
OP       IF cAction = 0
OP         LET(cAction, LASTCOL($nde, 1) + 1)
OP         SET_VALUE(CELL($nde, 1, cAction), "Action")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($nde, cSeam) var=r
OP         IF SAME(CELL($nde, r, cResult), "Fail")
OP           SET_VALUE(CELL($nde, r, cAction), "Repair")
OP           LET(n, n + 1)
OP         ELSE
OP           CLEAR(CELL($nde, r, cAction), what=contents)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Repairs sheet as a list, under headings, of the weld seam number and NDE method of every failed NDE report.
PROC     ListRepairs
KIND     Sub
MODULE   modSteps
ARGS     nde As Worksheet, repairs As Worksheet
VAR      cSeam As Long
VAR      cMethod As Long
VAR      cResult As Long
VAR      outRow As Long
PRE      COL($nde, "Weld Seam No") > 0; COL($nde, "Method") > 0; COL($nde, "Result") > 0
MODIFIES $repairs
OP       LET(cSeam, COL($nde, "Weld Seam No"))
OP       LET(cMethod, COL($nde, "Method"))
OP       LET(cResult, COL($nde, "Result"))
OP       CLEAR(USED_RANGE($repairs), what=all)
OP       WRITE_ROW($repairs, 1, "Weld Seam No", "Method")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($nde, cSeam) var=r
OP         IF SAME(CELL($nde, r, cResult), "Fail")
OP           WRITE_ROW($repairs, outRow, CELL($nde, r, cSeam), CELL($nde, r, cMethod))
OP           LET(outRow, outRow + 1)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write each weld seam's NDE result from the NDE Report into the NDE Result column of the Weld Log, or leave it blank when the seam has no NDE row.
PROC     FillNdeResults
KIND     Sub
MODULE   modSteps
ARGS     nde As Worksheet, weldLog As Worksheet
VAR      cSeam As Long
VAR      cResult As Long
VAR      cLogSeam As Long
VAR      cOut As Long
VAR      results As Dictionary
VAR      got As String
PRE      COL($nde, "Weld Seam No") > 0; COL($nde, "Result") > 0; COL($weldLog, "Weld Seam No") > 0
MODIFIES $weldLog
OP       LET(cSeam, COL($nde, "Weld Seam No"))
OP       LET(cResult, COL($nde, "Result"))
OP       LET(cLogSeam, COL($weldLog, "Weld Seam No"))
OP       LET(cOut, COL($weldLog, "NDE Result"))
OP       IF cOut = 0
OP         LET(cOut, LASTCOL($weldLog, 1) + 1)
OP         SET_VALUE(CELL($weldLog, 1, cOut), "NDE Result")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($nde, cSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($nde, r, cSeam))
OP         SKIP_IF IS_EMPTY(CELL($nde, r, cResult))
OP         DICT_ADD(results, TRIM(CELL($nde, r, cSeam)), TRIM(CELL($nde, r, cResult)), mode=overwrite)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($weldLog, cLogSeam) var=w
OP         CLEAR(CELL($weldLog, w, cOut), what=contents)
OP         IF EXISTS(results, TRIM(CELL($weldLog, w, cLogSeam)))
OP           DICT_GET(results, TRIM(CELL($weldLog, w, cLogSeam)), into=got)
OP           SET_VALUE(CELL($weldLog, w, cOut), got)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Method Summary sheet with one row for each NDE method giving its number of reports and its number of fails.
PROC     WriteMethodSummary
KIND     Sub
MODULE   modSteps
ARGS     nde As Worksheet, summary As Worksheet
VAR      cSeam As Long
VAR      cMethod As Long
VAR      cResult As Long
VAR      reports As Dictionary
VAR      fails As Dictionary
VAR      nReports As Long
VAR      nFails As Long
VAR      outRow As Long
PRE      COL($nde, "Weld Seam No") > 0; COL($nde, "Method") > 0; COL($nde, "Result") > 0
MODIFIES $summary
OP       LET(cSeam, COL($nde, "Weld Seam No"))
OP       LET(cMethod, COL($nde, "Method"))
OP       LET(cResult, COL($nde, "Result"))
OP       CLEAR(USED_RANGE($summary), what=all)
OP       WRITE_ROW($summary, 1, "Method", "Reports", "Fails")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($nde, cSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($nde, r, cMethod))
OP         DICT_ADD(reports, TRIM(CELL($nde, r, cMethod)), 1, mode=sum)
OP         DICT_ADD(fails, TRIM(CELL($nde, r, cMethod)), IIF(SAME(CELL($nde, r, cResult), "Fail"), 1, 0), mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(reports)
OP         DICT_GET(reports, k, into=nReports)
OP         DICT_GET(fails, k, into=nFails)
OP         WRITE_ROW($summary, outRow, k, nReports, nFails)
OP         LET(outRow, outRow + 1)
ERR      standard
