VCL 0.6
BUILD
# ASSUME: the real sheets are "WPS Register", "PQR List", "Problem WPS" and "Process Summary"; all four exist, and the two output sheets hold nothing else
# ASSUME: WPS Register headers (row 1) include "WPS No", "Process", "PQR Ref" and "Remark"; PQR List headers include "PQR No" and "Result"
# ASSUME: PQR numbers match ignoring case, spaces at the ends and doubled spaces inside; "Fail" is matched ignoring case and extra spaces
# ASSUME: a PQR No listed more than once counts as failed when any of its rows says Fail; a PQR with a blank or other Result is in the list and not failed
# ASSUME: a row with a blank WPS No is not a WPS: its Remark is left alone and it is not listed or counted
# ASSUME: clearing a Remark removes whatever text it held; a WPS with a blank Process is counted under "(blank)"
# ASSUME: the TESTs use the sheets Register, PqrList, Problems and Summary
GOAL     Mark each WPS whose PQR failed or is missing, list those WPS on the Problem WPS sheet and rebuild a summary of WPS and remarks by process.
NAME     PqrFollowUp
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(reg As Worksheet, pqr As Worksheet, problems As Worksheet, summary As Worksheet) As Long
SIG      modSteps.MarkPqrRemarks(reg As Worksheet, pqr As Worksheet) As Long
SIG      modSteps.ListProblemWps(reg As Worksheet, out As Worksheet)
SIG      modSteps.WriteProcessSummary(reg As Worksheet, out As Worksheet)
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.MarkPqrRemarks
FLOW     modMain.RunHeadless -> modSteps.ListProblemWps
FLOW     modMain.RunHeadless -> modSteps.WriteProcessSummary
FIXTURE  pqr_followup_01.xlsx sheet=Register cols=A:WPS No,B:Process,C:PQR Ref,D:Remark rows=11 data=A:values("W-07"x1,"W-03"x1,"W-05"x1,"W-01"x1,"W-06"x1,"W-02"x1,""x1,"W-04"x1,"W-09"x1,"W-10"x1,"W-08"x1)+boundary("W-03","W-07","W-05","W-01","W-06","W-02","","W-04","W-09","W-10","W-08");B:values("GTAW"x3," GTAW "x1,"SMAW"x2,"SAW"x2,""x1,"smaw"x1,"FCAW"x1)+boundary("SMAW","GTAW"," GTAW ","SAW","SMAW","GTAW","GTAW","","smaw","FCAW","SAW");C:values("PQR-01"x1,"pqr-02 "x1,"PQR-03"x1,"PQR-06"x1,"PQR-99"x1,"PQR-04"x1,"PQR-02"x1," PQR-05"x1,"PQR-07"x1,"PQR No"x1,""x1)+boundary("pqr-02 ","PQR-01","PQR-03","PQR-06","PQR-99","PQR-04","PQR-02"," PQR-05","PQR-07","PQR No","");D:values("old note"x1,""x6,"PQR failed"x1,"keep"x1,"PQR missing"x2)+boundary("","old note","","PQR failed","","","keep","PQR missing","PQR missing","","")
FIXTURE  pqr_followup_01.xlsx sheet=PqrList cols=A:PQR No,B:Process,C:Test Date,D:Result rows=10 data=A:values("PQR-01"x1,"PQR-02"x1," pqr-03 "x1,"PQR-04"x2,"PQR-05"x2,""x1,"PQR-07"x1,"PQR-06"x1)+boundary("PQR-01","PQR-02"," pqr-03 ","PQR-04","PQR-04","PQR-05","PQR-05","","PQR-07","PQR-06");B:values("GTAW"x4,"SMAW"x3,"SAW"x3);C:dates(2026-01-01,2026-09-30);D:values("Pass"x4,"Fail"x4,"fail"x1,""x1)+boundary("Pass","Fail","fail","Pass","Fail","Fail","Pass","Fail","Pass","")
FIXTURE  pqr_followup_01.xlsx sheet=Problems cols=A:WPS No,B:Process,C:PQR Ref,D:Remark rows=0 stale=10
FIXTURE  pqr_followup_01.xlsx sheet=Summary cols=A:Process,B:WPS,C:With Remark rows=0 stale=7
TEST     fixture=pqr_followup_01.xlsx call=modMain.RunHeadless(Register, PqrList, Problems, Summary) runs=2 assert=RESULT = 7 AND ERROR_COUNT = 0 AND COUNTIF(Register![Remark], "PQR failed") = 4 AND COUNTIF(Register![Remark], "PQR missing") = 3 AND COUNTA(Register![Remark]) = 9 AND VALUE(Register![Remark]2) = "PQR failed" AND VALUE(Register![Remark]4) = "PQR failed" AND VALUE(Register![Remark]7) = "PQR failed" AND VALUE(Register![Remark]9) = "PQR failed" AND VALUE(Register![Remark]6) = "PQR missing" AND VALUE(Register![Remark]11) = "PQR missing" AND VALUE(Register![Remark]12) = "PQR missing" AND VALUE(Register![Remark]8) = "keep" AND ROWS_USED(Problems) = 8 AND ROWS_USED(Summary) = 6 AND COUNTIF(Problems![WPS No], "stale") = 0 AND COUNTIF(Summary![Process], "stale") = 0 AND FILL_COUNT(Problems!A1:D12, #FF0000) = 0 AND FILL_COUNT(Summary!A1:C9, #FF0000) = 0
TEST     fixture=pqr_followup_01.xlsx call=modMain.RunHeadless(Register, PqrList, Problems, Summary) assert=ERROR_COUNT = 0 AND VALUE(Problems![WPS No]2) = "W-02" AND VALUE(Problems![Process]2) = "GTAW" AND VALUE(Problems![PQR Ref]2) = "PQR-04" AND VALUE(Problems![Remark]2) = "PQR failed" AND VALUE(Problems![WPS No]3) = "W-03" AND VALUE(Problems![WPS No]4) = "W-04" AND VALUE(Problems![WPS No]5) = "W-05" AND VALUE(Problems![WPS No]6) = "W-06" AND VALUE(Problems![WPS No]7) = "W-08" AND VALUE(Problems![WPS No]8) = "W-10" AND VALUE(Problems![Process]8) = "FCAW" AND VALUE(Problems![PQR Ref]8) = "PQR No" AND VALUE(Problems![Remark]8) = "PQR missing" AND COUNTIF(Problems![Remark], "PQR failed") = 4 AND COUNTIF(Problems![Remark], "PQR missing") = 3
TEST     fixture=pqr_followup_01.xlsx call=modMain.RunHeadless(Register, PqrList, Problems, Summary) assert=RESULT = 7 AND ERROR_COUNT = 0 AND VALUE(Summary![Process]2) = "(blank)" AND VALUE(Summary![WPS]2) = 1 AND VALUE(Summary![With Remark]2) = 1 AND VALUE(Summary![Process]3) = "FCAW" AND VALUE(Summary![WPS]3) = 1 AND VALUE(Summary![With Remark]3) = 1 AND VALUE(Summary![Process]4) = "GTAW" AND VALUE(Summary![WPS]4) = 3 AND VALUE(Summary![With Remark]4) = 2 AND VALUE(Summary![Process]5) = "SAW" AND VALUE(Summary![WPS]5) = 2 AND VALUE(Summary![With Remark]5) = 1 AND VALUE(Summary![Process]6) = "SMAW" AND VALUE(Summary![WPS]6) = 3 AND VALUE(Summary![With Remark]6) = 2 AND COUNTA(Register![WPS No]) = PRE_COUNTA(Register![WPS No]) AND COUNTA(Register![PQR Ref]) = PRE_COUNTA(Register![PQR Ref]) AND COUNTA(PqrList![PQR No]) = PRE_COUNTA(PqrList![PQR No]) AND COUNTIF(PqrList![Result], "Fail") = PRE_COUNTIF(PqrList![Result], "Fail")
TEST     fixture=pqr_followup_01.xlsx call=modMain.RunHeadless(PqrList, Register, Problems, Summary) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
TASK
GOAL     Start the PQR follow-up on the workbook's own sheets and tell the user how many WPS have a remark.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "WPS Register"), SHEET(THIS_BOOK, "PQR List"), SHEET(THIS_BOOK, "Problem WPS"), SHEET(THIS_BOOK, "Process Summary"), into=n)
OP       IF n < 0
OP         MSG("The PQR follow-up did not run: check the sheet headings (see ForgeLog).")
OP       ELSE
OP         MSG("WPS with a remark: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run the three steps of the PQR follow-up on the four sheets it is given and return how many WPS have a remark.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     reg As Worksheet, pqr As Worksheet, problems As Worksheet, summary As Worksheet
RET      Long
VAR      marked As Long
PRE      COL($reg, "WPS No") > 0; COL($reg, "Process") > 0; COL($reg, "PQR Ref") > 0; COL($reg, "Remark") > 0; COL($pqr, "PQR No") > 0; COL($pqr, "Result") > 0
MODIFIES $reg["Remark"]
MODIFIES $problems
MODIFIES $summary
OP       CALL(modSteps.MarkPqrRemarks, $reg, $pqr, into=marked)
OP       CALL(modSteps.ListProblemWps, $reg, $problems)
OP       CALL(modSteps.WriteProcessSummary, $reg, $summary)
OP       RETURN(marked)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write PQR failed or PQR missing in the Remark of each WPS whose PQR failed or is not in the PQR list, clear the other Remarks and return how many WPS got a remark.
PROC     MarkPqrRemarks
KIND     Function
MODULE   modSteps
ARGS     reg As Worksheet, pqr As Worksheet
RET      Long
VAR      cWps As Long
VAR      cRef As Long
VAR      cRemark As Long
VAR      cNo As Long
VAR      cResult As Long
VAR      pqrs As Dictionary
VAR      ref As String
VAR      verdict As String
VAR      n As Long
PRE      COL($reg, "WPS No") > 0; COL($reg, "PQR Ref") > 0; COL($reg, "Remark") > 0; COL($pqr, "PQR No") > 0; COL($pqr, "Result") > 0
MODIFIES $reg["Remark"]
OP       LET(cWps, COL($reg, "WPS No"))
OP       LET(cRef, COL($reg, "PQR Ref"))
OP       LET(cRemark, COL($reg, "Remark"))
OP       LET(cNo, COL($pqr, "PQR No"))
OP       LET(cResult, COL($pqr, "Result"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($pqr, cNo) var=p
OP         SKIP_IF IS_EMPTY(CELL($pqr, p, cNo))
OP         IF SAME(CELL($pqr, p, cResult), "Fail")
OP           DICT_ADD(pqrs, CLEAN_TEXT(CELL($pqr, p, cNo)), "Fail", mode=overwrite)
OP         ELSE
OP           DICT_ADD(pqrs, CLEAN_TEXT(CELL($pqr, p, cNo)), "OK", mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($reg, cWps) var=r
OP         SKIP_IF IS_EMPTY(CELL($reg, r, cWps))
OP         LET(ref, CLEAN_TEXT(CELL($reg, r, cRef)))
OP         IF EXISTS(pqrs, ref)
OP           DICT_GET(pqrs, ref, into=verdict)
OP           IF verdict = "Fail"
OP             SET_VALUE(CELL($reg, r, cRemark), "PQR failed")
OP             LET(n, n + 1)
OP           ELSE
OP             CLEAR(CELL($reg, r, cRemark), what=contents)
OP         ELSE
OP           SET_VALUE(CELL($reg, r, cRemark), "PQR missing")
OP           LET(n, n + 1)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the problem sheet with every WPS that has a remark, under the register's four headings, sorted by WPS No from A to Z.
PROC     ListProblemWps
KIND     Sub
MODULE   modSteps
ARGS     reg As Worksheet, out As Worksheet
VAR      cWps As Long
VAR      cProc As Long
VAR      cRef As Long
VAR      cRemark As Long
VAR      outRow As Long
PRE      COL($reg, "WPS No") > 0; COL($reg, "Process") > 0; COL($reg, "PQR Ref") > 0; COL($reg, "Remark") > 0
MODIFIES $out
OP       LET(cWps, COL($reg, "WPS No"))
OP       LET(cProc, COL($reg, "Process"))
OP       LET(cRef, COL($reg, "PQR Ref"))
OP       LET(cRemark, COL($reg, "Remark"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "WPS No", "Process", "PQR Ref", "Remark")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($reg, cWps) var=r
OP         SKIP_IF IS_EMPTY(CELL($reg, r, cWps))
OP         SKIP_IF IS_EMPTY(CELL($reg, r, cRemark))
OP         WRITE_ROW($out, outRow, CELL($reg, r, cWps), CELL($reg, r, cProc), CELL($reg, r, cRef), CELL($reg, r, cRemark))
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the summary sheet with one row per process in the register, its number of WPS and how many of them have a remark, sorted A to Z.
PROC     WriteProcessSummary
KIND     Sub
MODULE   modSteps
ARGS     reg As Worksheet, out As Worksheet
VAR      cWps As Long
VAR      cProc As Long
VAR      cRemark As Long
VAR      counts As Dictionary
VAR      remarks As Dictionary
VAR      procName As String
VAR      hasRemark As Long
VAR      got As Long
VAR      gotRemark As Long
VAR      outRow As Long
PRE      COL($reg, "WPS No") > 0; COL($reg, "Process") > 0; COL($reg, "Remark") > 0
MODIFIES $out
OP       LET(cWps, COL($reg, "WPS No"))
OP       LET(cProc, COL($reg, "Process"))
OP       LET(cRemark, COL($reg, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($reg, cWps) var=r
OP         SKIP_IF IS_EMPTY(CELL($reg, r, cWps))
OP         LET(procName, IIF(IS_EMPTY(CELL($reg, r, cProc)), "(blank)", TRIM(CELL($reg, r, cProc))))
OP         LET(hasRemark, IIF(IS_EMPTY(CELL($reg, r, cRemark)), 0, 1))
OP         DICT_ADD(counts, procName, 1, mode=sum)
OP         DICT_ADD(remarks, procName, hasRemark, mode=sum)
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Process", "WPS", "With Remark")
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         DICT_GET(remarks, k, into=gotRemark)
OP         WRITE_ROW($out, outRow, k, got, gotRemark)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
ERR      standard
