VCL 0.6
BUILD
# ASSUME: WPS Register headers (row 1) are "WPS No", "Process", "PQR Ref", "Review Date" and "Remark"; Welder Quals headers are "Welder ID", "Process" and "Expiry Date"
# ASSUME: the real sheets are "WPS Register", "Welder Quals", "No WPS", "Process Summary" and "WPS Summary" and all five exist; the TESTs use Wps, Quals, NoWps, ProcSummary and WpsSummary
# ASSUME: a WPS is a register row with a WPS No; a row with a blank WPS No is skipped everywhere; a qualification is a Welder Quals row with a Welder ID
# ASSUME: a WPS is Check when its PQR Ref is blank or its Review Date is a real date more than 365 days before today (exactly 365 days is not); a blank or non-date Review Date is not old by itself
# ASSUME: Remark holds only the Check flag, so every other WPS gets an empty Remark
# ASSUME: processes are matched after TRIM, ignoring case, and shown as first typed; a WPS with a blank Process is in the total but has no Process Summary row
# ASSUME: a qualification with a blank Process counts as having no WPS and is listed on No WPS, but belongs to no process row
# ASSUME: the output sheets are cleared and rewritten on every run; Process Summary is sorted A to Z and WPS Summary has a Measure and a Value column
GOAL     Flag WPS that lack a PQR or are overdue for review, then rebuild the No WPS list, the Process Summary and the WPS Summary sheets.
NAME     WpsCheck
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modChecks
MODULE   modLists
MODULE   modSummary
SIG      modMain.Run()
SIG      modMain.RunHeadless(wps As Worksheet, quals As Worksheet, noWps As Worksheet, procSum As Worksheet, wpsSum As Worksheet) As Long
SIG      modChecks.MarkCheck(wps As Worksheet) As Long
SIG      modLists.ListNoWps(wps As Worksheet, quals As Worksheet, out As Worksheet) As Long
SIG      modSummary.WriteProcessSummary(wps As Worksheet, quals As Worksheet, out As Worksheet)
SIG      modSummary.WriteTotals(wps As Worksheet, out As Worksheet, flagged As Long, unmatched As Long)
FLOW     modMain.Run -> modMain.RunHeadless -> modChecks.MarkCheck
FLOW     modMain.RunHeadless -> modLists.ListNoWps
FLOW     modMain.RunHeadless -> modSummary.WriteProcessSummary
FLOW     modMain.RunHeadless -> modSummary.WriteTotals
FIXTURE  wps_01.xlsx sheet=Wps cols=A:WPS No,B:Process,C:PQR Ref,D:Review Date,E:Remark rows=12 data=A:values("WPS-01"x1,"WPS-02"x1,"WPS-03"x1,"WPS-04"x1,"WPS-05"x1,"WPS-06"x1,"WPS-07"x1,""x1,"WPS-09"x1,"WPS-10"x1,"WPS-11"x1,"WPS-12"x1)+boundary("WPS-01","WPS-02","WPS-03","WPS-04","WPS-05","WPS-06","WPS-07","","WPS-09","WPS-10","WPS-11","WPS-12");B:values("FCAW"x1,"GTAW"x3,"gtaw"x1,"SMAW"x2,"SAW"x1,""x1,"  GMAW "x1,"gmaw"x1,"GMAW"x1)+boundary("FCAW","GTAW","gtaw","SMAW","GTAW","GTAW","SMAW","SAW","","  GMAW ","gmaw","GMAW");C:values(""x3,"PQR-2"x1,"PQR-3"x1,"PQR-4"x1,"PQR-5"x1,"PQR-7"x1,"PQR-9"x1,"PQR-10"x1,"PQR-11"x1,"PQR-12"x1)+boundary("","PQR-2","PQR-3","PQR-4","PQR-5","","PQR-7","","PQR-9","PQR-10","PQR-11","PQR-12");D:reldates(-400,0)+boundary(-10,-365,-366,0,"","",-400,-500,-20,-30,-1,-2);E:values(""x8,"Check"x3,"Note"x1)+boundary("","Check","","","Check","","","Note","","","Check","")
FIXTURE  wps_01.xlsx sheet=Quals cols=A:Welder ID,B:Process,C:Expiry Date rows=13 data=A:values("W-01"x1,"W-02"x1,"W-03"x1,"W-04"x1,"W-05"x1,"W-06"x1,""x2,"W-09"x1,"W-10"x1,"W-11"x1,"W-12"x1,"W-13"x1)+boundary("W-01","W-02","W-03","W-04","W-05","W-06","","","W-09","W-10","W-11","W-12","W-13");B:values("FCAW"x1,"smaw"x1,"Smaw "x1,"PAW"x2,"paw"x1,"SAW"x1,"SMAW"x2,""x1,"GMAW"x1,"gmaw"x1,"TIG"x1)+boundary("FCAW","smaw","Smaw ","PAW","paw","SAW","PAW","SMAW","","GMAW","gmaw","SMAW","TIG");C:dates(2026-01-01,2027-12-31)
FIXTURE  wps_01.xlsx sheet=NoWps cols=A:Welder ID,B:Process,C:Expiry Date rows=0 stale=2
FIXTURE  wps_01.xlsx sheet=ProcSummary cols=A:Process,B:WPS,C:Check WPS,D:Qualifications rows=0 stale=2
FIXTURE  wps_01.xlsx sheet=WpsSummary cols=A:Measure,B:Value rows=0 stale=2
FIXTURE  wps_02.xlsx sheet=Wps cols=A:WPS No,B:Process,C:PQR Ref,D:Review Date,E:Remark rows=2 data=A:values("WPS-A1"x1,"WPS-A2"x1)+boundary("WPS-A1","WPS-A2");B:values("FCAW"x1,"GMAW"x1)+boundary("FCAW","GMAW");C:values("PQR-A1"x1,""x1)+boundary("PQR-A1","");D:reldates(-9,-1)+boundary(-5,-6);E:values(""x2)
FIXTURE  wps_02.xlsx sheet=Quals cols=A:Welder ID,B:Process,C:Expiry Date rows=2 data=A:values("W-A1"x1,"W-A2"x1)+boundary("W-A1","W-A2");B:values("PAW"x1,"FCAW"x1)+boundary("PAW","FCAW");C:dates(2026-01-01,2027-12-31)
FIXTURE  wps_02.xlsx sheet=NoWps cols=A:Welder ID,B:Process,C:Expiry Date rows=0 stale=1
FIXTURE  wps_02.xlsx sheet=ProcSummary cols=A:Process,B:WPS,C:Check WPS,D:Qualifications rows=0 stale=1
FIXTURE  wps_02.xlsx sheet=WpsSummary cols=A:Measure,B:Value rows=0 stale=1
TEST     fixture=wps_01.xlsx call=modMain.RunHeadless(Wps, Quals, NoWps, ProcSummary, WpsSummary) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(Wps![Remark], "Check") = 4 AND COUNTA(Wps![Remark]) = 6 AND ROWS_USED(NoWps) = 6 AND ROWS_USED(ProcSummary) = 5 AND ROWS_USED(WpsSummary) = 4 AND COUNTIF(NoWps![Welder ID], "stale") = 0 AND COUNTIF(ProcSummary![Process], "stale") = 0 AND COUNTIF(WpsSummary![Measure], "stale") = 0
TEST     fixture=wps_01.xlsx call=modMain.RunHeadless(Wps, Quals, NoWps, ProcSummary, WpsSummary) assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIFS(Wps![WPS No], "WPS-01", Wps![Remark], "Check") = 1 AND COUNTIFS(Wps![WPS No], "WPS-02", Wps![Remark], "Check") = 0 AND COUNTIFS(Wps![WPS No], "WPS-03", Wps![Remark], "Check") = 1 AND COUNTIFS(Wps![WPS No], "WPS-04", Wps![Remark], "Check") = 0 AND COUNTIFS(Wps![WPS No], "WPS-05", Wps![Remark], "Check") = 0 AND COUNTIFS(Wps![WPS No], "WPS-06", Wps![Remark], "Check") = 1 AND COUNTIFS(Wps![WPS No], "WPS-07", Wps![Remark], "Check") = 1 AND COUNTIFS(Wps![WPS No], "WPS-11", Wps![Remark], "Check") = 0 AND COUNTIF(Wps![Remark], "Note") = 1
TEST     fixture=wps_01.xlsx call=modMain.RunHeadless(Wps, Quals, NoWps, ProcSummary, WpsSummary) assert=ERROR_COUNT = 0 AND VALUE(NoWps![Welder ID]2) = "W-04" AND VALUE(NoWps![Process]2) = "PAW" AND VALUE(NoWps![Welder ID]3) = "W-05" AND VALUE(NoWps![Process]3) = "paw" AND VALUE(NoWps![Welder ID]4) = "W-06" AND VALUE(NoWps![Welder ID]5) = "W-09" AND VALUE(NoWps![Welder ID]6) = "W-13" AND VALUE(NoWps![Process]6) = "TIG" AND COUNTIF(NoWps![Welder ID], "W-02") = 0 AND COUNTIF(NoWps![Welder ID], "W-03") = 0 AND COUNTIF(NoWps![Welder ID], "W-10") = 0 AND COUNTA(NoWps![Expiry Date]) = 6
TEST     fixture=wps_01.xlsx call=modMain.RunHeadless(Wps, Quals, NoWps, ProcSummary, WpsSummary) assert=ERROR_COUNT = 0 AND VALUE(ProcSummary![Process]2) = "FCAW" AND VALUE(ProcSummary![WPS]2) = 1 AND VALUE(ProcSummary![Check WPS]2) = 1 AND VALUE(ProcSummary![Qualifications]2) = 1 AND VALUE(ProcSummary![Process]3) = "GMAW" AND VALUE(ProcSummary![WPS]3) = 3 AND VALUE(ProcSummary![Check WPS]3) = 0 AND VALUE(ProcSummary![Qualifications]3) = 2 AND VALUE(ProcSummary![Process]4) = "GTAW" AND VALUE(ProcSummary![WPS]4) = 4 AND VALUE(ProcSummary![Check WPS]4) = 2 AND VALUE(ProcSummary![Qualifications]4) = 0 AND VALUE(ProcSummary![Process]5) = "SMAW" AND VALUE(ProcSummary![WPS]5) = 2 AND VALUE(ProcSummary![Check WPS]5) = 1 AND VALUE(ProcSummary![Qualifications]5) = 3
TEST     fixture=wps_01.xlsx call=modMain.RunHeadless(Wps, Quals, NoWps, ProcSummary, WpsSummary) assert=ERROR_COUNT = 0 AND VALUE(WpsSummary![Measure]2) = "Total WPS" AND VALUE(WpsSummary![Value]2) = 11 AND VALUE(WpsSummary![Measure]3) = "Check WPS" AND VALUE(WpsSummary![Value]3) = 4 AND VALUE(WpsSummary![Measure]4) = "Qualifications with no WPS" AND VALUE(WpsSummary![Value]4) = 5
TEST     fixture=wps_01.xlsx call=modMain.RunHeadless(Wps, Quals, NoWps, ProcSummary, WpsSummary) assert=ERROR_COUNT = 0 AND COUNTA(Wps![WPS No]) = PRE_COUNTA(Wps![WPS No]) AND COUNTA(Wps![Process]) = PRE_COUNTA(Wps![Process]) AND COUNTA(Wps![PQR Ref]) = PRE_COUNTA(Wps![PQR Ref]) AND COUNTA(Wps![Review Date]) = PRE_COUNTA(Wps![Review Date]) AND COUNTA(Quals![Welder ID]) = PRE_COUNTA(Quals![Welder ID]) AND COUNTA(Quals![Process]) = PRE_COUNTA(Quals![Process]) AND COUNTA(Quals![Expiry Date]) = PRE_COUNTA(Quals![Expiry Date])
TEST     fixture=wps_02.xlsx call=modMain.RunHeadless(Wps, Quals, NoWps, ProcSummary, WpsSummary) assert=RESULT = 1 AND ERROR_COUNT = 0 AND COUNTA(Wps![Remark]) = 2 AND ROWS_USED(NoWps) = 2 AND VALUE(NoWps![Welder ID]2) = "W-A1" AND VALUE(ProcSummary![Process]2) = "FCAW" AND VALUE(ProcSummary![Check WPS]2) = 0 AND VALUE(ProcSummary![Qualifications]2) = 1 AND VALUE(ProcSummary![Process]3) = "GMAW" AND VALUE(ProcSummary![Check WPS]3) = 1 AND VALUE(ProcSummary![Qualifications]3) = 0 AND VALUE(WpsSummary![Value]2) = 2 AND VALUE(WpsSummary![Value]3) = 1 AND VALUE(WpsSummary![Value]4) = 1
TEST     fixture=wps_01.xlsx call=modMain.RunHeadless(Quals, Wps, NoWps, ProcSummary, WpsSummary) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
TASK
GOAL     Start the WPS check on the workbook's own sheets and tell the user how many WPS were marked Check.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "WPS Register"), SHEET(THIS_BOOK, "Welder Quals"), SHEET(THIS_BOOK, "No WPS"), SHEET(THIS_BOOK, "Process Summary"), SHEET(THIS_BOOK, "WPS Summary"), into=n)
OP       MSG("WPS marked Check: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run every step of the WPS check on the five sheets it is given and return how many WPS were marked Check.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     wps As Worksheet, quals As Worksheet, noWps As Worksheet, procSum As Worksheet, wpsSum As Worksheet
RET      Long
VAR      flagged As Long
VAR      unmatched As Long
PRE      COL($wps, "WPS No") > 0; COL($wps, "Process") > 0; COL($wps, "PQR Ref") > 0; COL($wps, "Review Date") > 0; COL($wps, "Remark") > 0; COL($quals, "Welder ID") > 0; COL($quals, "Process") > 0; COL($quals, "Expiry Date") > 0
MODIFIES $wps["Remark"]
MODIFIES $noWps
MODIFIES $procSum
MODIFIES $wpsSum
OP       CALL(modChecks.MarkCheck, $wps, into=flagged)
OP       CALL(modLists.ListNoWps, $wps, $quals, $noWps, into=unmatched)
OP       CALL(modSummary.WriteProcessSummary, $wps, $quals, $procSum)
OP       CALL(modSummary.WriteTotals, $wps, $wpsSum, flagged, unmatched)
OP       RETURN(flagged)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Check in Remark for every WPS with no PQR reference or a review date over 365 days old, and clear Remark on the others.
PROC     MarkCheck
KIND     Function
MODULE   modChecks
ARGS     wps As Worksheet
RET      Long
VAR      cNo As Long
VAR      cPqr As Long
VAR      cDate As Long
VAR      cRemark As Long
VAR      old As Boolean
VAR      flagged As Long
PRE      COL($wps, "WPS No") > 0; COL($wps, "PQR Ref") > 0; COL($wps, "Review Date") > 0; COL($wps, "Remark") > 0
MODIFIES $wps["Remark"]
OP       LET(cNo, COL($wps, "WPS No"))
OP       LET(cPqr, COL($wps, "PQR Ref"))
OP       LET(cDate, COL($wps, "Review Date"))
OP       LET(cRemark, COL($wps, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($wps, cNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($wps, r, cNo))
OP         LET(old, FALSE)
OP         IF IS_DATE(CELL($wps, r, cDate))
OP           IF CELL($wps, r, cDate) < TODAY - 365
OP             LET(old, TRUE)
OP         IF IS_EMPTY(CELL($wps, r, cPqr)) OR old
OP           SET_VALUE(CELL($wps, r, cRemark), "Check")
OP           LET(flagged, flagged + 1)
OP         ELSE
OP           CLEAR(CELL($wps, r, cRemark), what=contents)
OP       RETURN(flagged)
ERR      standard
---
VCL 0.6
TASK
GOAL     List every welder qualification whose process has no WPS in the register on the output sheet and return how many were listed.
PROC     ListNoWps
KIND     Function
MODULE   modLists
ARGS     wps As Worksheet, quals As Worksheet, out As Worksheet
RET      Long
VAR      cNo As Long
VAR      cProc As Long
VAR      cWelder As Long
VAR      cQProc As Long
VAR      cExpiry As Long
VAR      known As Dictionary
VAR      outRow As Long
VAR      listed As Long
PRE      COL($wps, "WPS No") > 0; COL($wps, "Process") > 0; COL($quals, "Welder ID") > 0; COL($quals, "Process") > 0; COL($quals, "Expiry Date") > 0
MODIFIES $out
OP       LET(cNo, COL($wps, "WPS No"))
OP       LET(cProc, COL($wps, "Process"))
OP       LET(cWelder, COL($quals, "Welder ID"))
OP       LET(cQProc, COL($quals, "Process"))
OP       LET(cExpiry, COL($quals, "Expiry Date"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Welder ID", "Process", "Expiry Date")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($wps, cNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($wps, r, cNo)) OR IS_EMPTY(CELL($wps, r, cProc))
OP         DICT_ADD(known, TRIM(CELL($wps, r, cProc)), 1, mode=keep)
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($quals, cWelder) var=q
OP         SKIP_IF IS_EMPTY(CELL($quals, q, cWelder))
OP         IF NOT EXISTS(known, TRIM(CELL($quals, q, cQProc)))
OP           WRITE_ROW($out, outRow, CELL($quals, q, cWelder), CELL($quals, q, cQProc), CELL($quals, q, cExpiry))
OP           LET(outRow, outRow + 1)
OP           LET(listed, listed + 1)
OP       RETURN(listed)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the process summary with each register process, its WPS count, its Check count and its qualification count, A to Z.
PROC     WriteProcessSummary
KIND     Sub
MODULE   modSummary
ARGS     wps As Worksheet, quals As Worksheet, out As Worksheet
VAR      cNo As Long
VAR      cProc As Long
VAR      cRemark As Long
VAR      cWelder As Long
VAR      cQProc As Long
VAR      wpsCounts As Dictionary
VAR      checkCounts As Dictionary
VAR      qualCounts As Dictionary
VAR      nWps As Long
VAR      nCheck As Long
VAR      nQual As Long
VAR      outRow As Long
PRE      COL($wps, "WPS No") > 0; COL($wps, "Process") > 0; COL($wps, "Remark") > 0; COL($quals, "Welder ID") > 0; COL($quals, "Process") > 0
MODIFIES $out
OP       LET(cNo, COL($wps, "WPS No"))
OP       LET(cProc, COL($wps, "Process"))
OP       LET(cRemark, COL($wps, "Remark"))
OP       LET(cWelder, COL($quals, "Welder ID"))
OP       LET(cQProc, COL($quals, "Process"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Process", "WPS", "Check WPS", "Qualifications")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($wps, cNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($wps, r, cNo)) OR IS_EMPTY(CELL($wps, r, cProc))
OP         DICT_ADD(wpsCounts, TRIM(CELL($wps, r, cProc)), 1, mode=sum)
OP         IF SAME(CELL($wps, r, cRemark), "Check")
OP           DICT_ADD(checkCounts, TRIM(CELL($wps, r, cProc)), 1, mode=sum)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($quals, cWelder) var=q
OP         SKIP_IF IS_EMPTY(CELL($quals, q, cWelder)) OR IS_EMPTY(CELL($quals, q, cQProc))
OP         DICT_ADD(qualCounts, TRIM(CELL($quals, q, cQProc)), 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(wpsCounts)
OP         DICT_GET(wpsCounts, k, into=nWps)
OP         LET(nCheck, 0)
OP         IF EXISTS(checkCounts, k) THEN DICT_GET(checkCounts, k, into=nCheck)
OP         LET(nQual, 0)
OP         IF EXISTS(qualCounts, k) THEN DICT_GET(qualCounts, k, into=nQual)
OP         WRITE_ROW($out, outRow, k, nWps, nCheck, nQual)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the WPS summary sheet with labelled totals for all WPS, Check WPS and qualifications that have no WPS.
PROC     WriteTotals
KIND     Sub
MODULE   modSummary
ARGS     wps As Worksheet, out As Worksheet, flagged As Long, unmatched As Long
VAR      cNo As Long
VAR      total As Long
PRE      COL($wps, "WPS No") > 0
MODIFIES $out
OP       LET(cNo, COL($wps, "WPS No"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($wps, cNo) var=r
OP         IF NOT IS_EMPTY(CELL($wps, r, cNo)) THEN LET(total, total + 1)
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Measure", "Value")
OP       WRITE_ROW($out, 2, "Total WPS", total)
OP       WRITE_ROW($out, 3, "Check WPS", $flagged)
OP       WRITE_ROW($out, 4, "Qualifications with no WPS", $unmatched)
ERR      standard
