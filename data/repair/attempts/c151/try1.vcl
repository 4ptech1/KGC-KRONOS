VCL 0.6
BUILD
# ASSUME: WPS Register headers include "WPS No", "Process", "PQR Ref" and "Remark"; PQR List has "PQR No" and "Result"; Welder Quals has "Welder ID", "Process" and "WPS Check"
# ASSUME: the real sheets are "WPS Register", "PQR List", "Welder Quals" and "Process Coverage"; the TESTs use Wps, Pqr, Welders and Cov
# ASSUME: processes, PQR numbers and IDs are matched ignoring case and extra spaces; a register or welder row with a blank WPS No or Welder ID is skipped, and so is a WPS or welder row with a blank Process in the coverage and in the welder check
# ASSUME: a PQR counts as failed when its Result is Fail, ignoring case and spaces; PQR numbers are unique in the PQR List
# ASSUME: a welder with a blank process has no WPS; a WPS has a remark when its Remark is not blank, after step 1 has run
# ASSUME: the sheet Process Coverage is rebuilt on every run and exists; its headings are Process, WPS Count, WPS With Remark and Welders Qualified
# ASSUME: the welders qualified for a process are the welder rows with that process
GOAL     Remark WPS without a good PQR, mark welders with no clean WPS for their process, rebuild a process coverage table and count the No WPS welders.
NAME     WpsCoverageCheck
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(wps As Worksheet, pqr As Worksheet, welders As Worksheet, out As Worksheet) As Long
SIG      modSteps.MarkWpsRemarks(wps As Worksheet, pqr As Worksheet)
SIG      modSteps.MarkWelderWps(welders As Worksheet, wps As Worksheet) As Long
SIG      modSteps.WriteCoverage(wps As Worksheet, welders As Worksheet, out As Worksheet)
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.MarkWpsRemarks
FLOW     modMain.RunHeadless -> modSteps.MarkWelderWps
FLOW     modMain.RunHeadless -> modSteps.WriteCoverage
FIXTURE  wps_01.xlsx sheet=Wps cols=A:WPS No,B:Process,C:PQR Ref,D:Remark rows=9 data=A:values("W1"x1,"W2"x1,"W3"x1,"W4"x1,"W5"x1,"W6"x1,"W7"x1,""x1,"W9"x1)+boundary("W1","W2","W3","W4","W5","W6","W7","","W9");B:values("GTAW"x2,"SMAW"x2," smaw "x1,"FCAW"x2,"gtaw"x1,""x1)+boundary("GTAW","GTAW","SMAW","SMAW"," smaw ","FCAW","gtaw","FCAW","");C:values("PQR-1"x1," pqr-2 "x1,""x2,"PQR-9"x1,"PQR-3"x2,"PQR-4"x1,"pqr-1"x1)+boundary("PQR-1"," pqr-2 ","","PQR-9","PQR-3","PQR-4","pqr-1","","PQR-3");D:values("old"x8,""x1)+boundary("old","old","old","old","old","old","old","","old")
FIXTURE  wps_01.xlsx sheet=Pqr cols=A:PQR No,B:Process,C:Result rows=5 data=A:values("PQR-1"x1,"PQR-2"x1,""x1,"PQR-3"x1,"PQR-4"x1)+boundary("PQR-1","PQR-2","","PQR-3","PQR-4");B:text(4);C:values("Pass"x3,"FAIL"x1," Fail "x1)+boundary("Pass","FAIL","Pass","Pass"," Fail ")
FIXTURE  wps_01.xlsx sheet=Welders cols=A:Welder ID,B:Process,C:WPS Check rows=8 data=A:values("A1"x1,"A2"x1,"A3"x1,"A4"x1,"A5"x1,"A6"x1,""x1,"A8"x1)+boundary("A1","A2","A3","A4","A5","A6","","A8");B:values("GTAW"x2," gtaw "x1,"SMAW"x1,"FCAW"x1,"MIG"x1,""x1,"smaw"x1)+boundary("GTAW"," gtaw ","SMAW","FCAW","MIG","","GTAW","smaw");C:values("old"x8)
FIXTURE  wps_01.xlsx sheet=Cov cols=A:Process,B:WPS Count,C:WPS With Remark,D:Welders Qualified rows=0 stale=2
TEST     fixture=wps_01.xlsx call=modMain.RunHeadless(Wps, Pqr, Welders, Cov) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Wps![Remark], "No PQR") = 2 AND COUNTIF(Wps![Remark], "PQR failed") = 2 AND COUNTIF(Wps![Remark], "old") = 0 AND COUNTA(Wps![Remark]) = 5 AND COUNTIF(Welders![WPS Check], "No WPS") = 3 AND COUNTIF(Welders![WPS Check], "old") = 1 AND COUNTA(Welders![WPS Check]) = 5 AND COUNTA(Cov![Process]) = 4 AND COUNTIF(Cov![Process], "stale") = 0
TEST     fixture=wps_01.xlsx call=modMain.RunHeadless(Wps, Pqr, Welders, Cov) assert=RESULT = 3 AND COUNTA(Wps![WPS No]) = PRE_COUNTA(Wps![WPS No]) AND COUNTA(Pqr![PQR No]) = PRE_COUNTA(Pqr![PQR No]) AND COUNTA(Welders![Welder ID]) = PRE_COUNTA(Welders![Welder ID]) AND VALUE(Cov![Process]2) = "FCAW" AND VALUE(Cov![WPS Count]2) = 1 AND VALUE(Cov![WPS With Remark]2) = 1 AND VALUE(Cov![Welders Qualified]2) = 1 AND VALUE(Cov![Process]3) = "GTAW" AND VALUE(Cov![WPS Count]3) = 3 AND VALUE(Cov![WPS With Remark]3) = 1 AND VALUE(Cov![Welders Qualified]3) = 2 AND VALUE(Cov![Process]4) = "SMAW" AND VALUE(Cov![WPS Count]4) = 3 AND VALUE(Cov![WPS With Remark]4) = 2 AND VALUE(Cov![Welders Qualified]4) = 2
TEST     fixture=wps_01.xlsx call=modMain.RunHeadless(Pqr, Wps, Welders, Cov) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
TASK
GOAL     Start the WPS coverage check on the workbook's own sheets and tell the user how many welder rows have No WPS.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "WPS Register"), SHEET(THIS_BOOK, "PQR List"), SHEET(THIS_BOOK, "Welder Quals"), SHEET(THIS_BOOK, "Process Coverage"), into=n)
OP       MSG("Welder rows with No WPS: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run all three steps on the four sheets it is given and return how many welder rows have No WPS.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     wps As Worksheet, pqr As Worksheet, welders As Worksheet, out As Worksheet
RET      Long
VAR      noWps As Long
PRE      COL($wps, "WPS No") > 0; COL($wps, "Process") > 0; COL($wps, "PQR Ref") > 0; COL($wps, "Remark") > 0; COL($pqr, "PQR No") > 0; COL($pqr, "Result") > 0; COL($welders, "Welder ID") > 0; COL($welders, "Process") > 0; COL($welders, "WPS Check") > 0
MODIFIES $wps["Remark"]
MODIFIES $welders["WPS Check"]
MODIFIES $out
OP       CALL(modSteps.MarkWpsRemarks, $wps, $pqr)
OP       CALL(modSteps.MarkWelderWps, $welders, $wps, into=noWps)
OP       CALL(modSteps.WriteCoverage, $wps, $welders, $out)
OP       RETURN(noWps)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write No PQR or PQR failed in the Remark of each WPS whose PQR is missing or failed, and clear the Remark of the others.
PROC     MarkWpsRemarks
KIND     Sub
MODULE   modSteps
ARGS     wps As Worksheet, pqr As Worksheet
VAR      cNo As Long
VAR      cRef As Long
VAR      cRemark As Long
VAR      cPqr As Long
VAR      cResult As Long
VAR      results As Dictionary
VAR      pqrResult As String
VAR      flag As String
PRE      COL($wps, "WPS No") > 0; COL($wps, "PQR Ref") > 0; COL($wps, "Remark") > 0; COL($pqr, "PQR No") > 0; COL($pqr, "Result") > 0
MODIFIES $wps["Remark"]
OP       LET(cNo, COL($wps, "WPS No"))
OP       LET(cRef, COL($wps, "PQR Ref"))
OP       LET(cRemark, COL($wps, "Remark"))
OP       LET(cPqr, COL($pqr, "PQR No"))
OP       LET(cResult, COL($pqr, "Result"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($pqr, cPqr) var=r
OP         SKIP_IF IS_EMPTY(CELL($pqr, r, cPqr))
OP         DICT_ADD(results, CLEAN_TEXT(CELL($pqr, r, cPqr)), CELL($pqr, r, cResult), mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($wps, cNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($wps, r, cNo))
OP         LET(flag, "")
OP         IF NOT EXISTS(results, CLEAN_TEXT(CELL($wps, r, cRef)))
OP           LET(flag, "No PQR")
OP         ELSE
OP           DICT_GET(results, CLEAN_TEXT(CELL($wps, r, cRef)), into=pqrResult)
OP           IF SAME(pqrResult, "Fail")
OP             LET(flag, "PQR failed")
OP         IF flag = ""
OP           CLEAR(CELL($wps, r, cRemark), what=contents)
OP         ELSE
OP           SET_VALUE(CELL($wps, r, cRemark), flag)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write No WPS in WPS Check for each welder whose process has no WPS without a remark, clear it for the others, and return how many got No WPS.
PROC     MarkWelderWps
KIND     Function
MODULE   modSteps
ARGS     welders As Worksheet, wps As Worksheet
RET      Long
VAR      cNo As Long
VAR      cProc As Long
VAR      cRemark As Long
VAR      cId As Long
VAR      cWProc As Long
VAR      cCheck As Long
VAR      clean As Dictionary
VAR      n As Long
PRE      COL($wps, "WPS No") > 0; COL($wps, "Process") > 0; COL($wps, "Remark") > 0; COL($welders, "Welder ID") > 0; COL($welders, "Process") > 0; COL($welders, "WPS Check") > 0
MODIFIES $welders["WPS Check"]
OP       LET(cNo, COL($wps, "WPS No"))
OP       LET(cProc, COL($wps, "Process"))
OP       LET(cRemark, COL($wps, "Remark"))
OP       LET(cId, COL($welders, "Welder ID"))
OP       LET(cWProc, COL($welders, "Process"))
OP       LET(cCheck, COL($welders, "WPS Check"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($wps, cNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($wps, r, cNo))
OP         SKIP_IF IS_EMPTY(CELL($wps, r, cProc))
OP         IF IS_EMPTY(CELL($wps, r, cRemark))
OP           DICT_ADD(clean, CLEAN_TEXT(CELL($wps, r, cProc)), 1, mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($welders, cId) var=r
OP         SKIP_IF IS_EMPTY(CELL($welders, r, cId))
OP         IF EXISTS(clean, CLEAN_TEXT(CELL($welders, r, cWProc)))
OP           CLEAR(CELL($welders, r, cCheck), what=contents)
OP         ELSE
OP           SET_VALUE(CELL($welders, r, cCheck), "No WPS")
OP           LET(n, n + 1)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Process Coverage sheet with each process of the WPS Register, its WPS count, WPS with a remark and qualified welders, A to Z.
PROC     WriteCoverage
KIND     Sub
MODULE   modSteps
ARGS     wps As Worksheet, welders As Worksheet, out As Worksheet
VAR      cNo As Long
VAR      cProc As Long
VAR      cRemark As Long
VAR      cId As Long
VAR      cWProc As Long
VAR      counts As Dictionary
VAR      remarked As Dictionary
VAR      qualified As Dictionary
VAR      a As Long
VAR      b As Long
VAR      c As Long
VAR      outRow As Long
PRE      COL($wps, "WPS No") > 0; COL($wps, "Process") > 0; COL($wps, "Remark") > 0; COL($welders, "Welder ID") > 0; COL($welders, "Process") > 0
MODIFIES $out
OP       LET(cNo, COL($wps, "WPS No"))
OP       LET(cProc, COL($wps, "Process"))
OP       LET(cRemark, COL($wps, "Remark"))
OP       LET(cId, COL($welders, "Welder ID"))
OP       LET(cWProc, COL($welders, "Process"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Process", "WPS Count", "WPS With Remark", "Welders Qualified")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($wps, cNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($wps, r, cNo))
OP         SKIP_IF IS_EMPTY(CELL($wps, r, cProc))
OP         DICT_ADD(counts, CLEAN_TEXT(CELL($wps, r, cProc)), 1, mode=sum)
OP         DICT_ADD(remarked, CLEAN_TEXT(CELL($wps, r, cProc)), IIF(IS_EMPTY(CELL($wps, r, cRemark)), 0, 1), mode=sum)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($welders, cId) var=r
OP         SKIP_IF IS_EMPTY(CELL($welders, r, cId))
OP         SKIP_IF IS_EMPTY(CELL($welders, r, cWProc))
OP         DICT_ADD(qualified, CLEAN_TEXT(CELL($welders, r, cWProc)), 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=a)
OP         DICT_GET(remarked, k, into=b)
OP         LET(c, 0)
OP         IF EXISTS(qualified, k)
OP           DICT_GET(qualified, k, into=c)
OP         WRITE_ROW($out, outRow, k, a, b, c)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
ERR      standard
