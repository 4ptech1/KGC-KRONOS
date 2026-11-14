VCL 0.6
TASK
# ASSUME: PQR List headers in row 1 include "Process" and "Result"; the PQR Summary sheet exists and holds nothing else
# ASSUME: every row with a Process counts as one PQR, even when its PQR No is blank or repeated
# ASSUME: processes are grouped ignoring case and spaces at the ends, and shown as first typed
# ASSUME: "Fail" is matched ignoring case and extra spaces; any other Result, a blank one too, is not a fail
GOAL     Rebuild the PQR Summary sheet with one row per welding process, its number of PQRs and how many of them failed, sorted A to Z, and a Total row at the end.
PROC     WritePqrSummary
KIND     Function
MODULE   modPqr
ARGS     src As Worksheet, out As Worksheet
RET      Long
VAR      cProc As Long
VAR      cResult As Long
VAR      counts As Dictionary
VAR      fails As Dictionary
VAR      isFail As Long
VAR      total As Long
VAR      totalFail As Long
VAR      got As Long
VAR      gotFail As Long
VAR      outRow As Long
PRE      COL($src, "Process") > 0; COL($src, "Result") > 0
MODIFIES $out
OP       LET(cProc, COL($src, "Process"))
OP       LET(cResult, COL($src, "Result"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($src, cProc) var=r
OP         SKIP_IF IS_EMPTY(CELL($src, r, cProc))
OP         LET(isFail, IIF(SAME(CELL($src, r, cResult), "Fail"), 1, 0))
OP         DICT_ADD(counts, TRIM(CELL($src, r, cProc)), 1, mode=sum)
OP         DICT_ADD(fails, TRIM(CELL($src, r, cProc)), isFail, mode=sum)
OP         LET(total, total + 1)
OP         LET(totalFail, totalFail + isFail)
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Process", "PQRs", "Failed")
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         DICT_GET(fails, k, into=gotFail)
OP         WRITE_ROW($out, outRow, k, got, gotFail)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       WRITE_ROW($out, outRow, "Total", total, totalFail)
OP       RETURN(COUNT(counts))
ERR      standard
FIXTURE  pqr_list_01.xlsx sheet=PqrList cols=A:PQR No,B:Process,C:Test Date,D:Result rows=12 data=A:seq(1,1);B:values("SMAW"x2,"GTAW"x3,""x2,"smaw"x1," GTAW "x1,"SAW"x2,"FCAW"x1)+boundary("SMAW","GTAW","","smaw"," GTAW ","SAW","GTAW","SMAW","FCAW","","GTAW","SAW");C:dates(2026-01-01,2026-09-30);D:values("Fail"x3,"Pass"x5,"FAIL"x1,"fail "x1,"Failed"x1,""x1)+boundary("Fail","Pass","Fail","FAIL","fail ","Pass","Pass","Pass","Fail","Pass","Failed","")
FIXTURE  pqr_list_01.xlsx sheet=Summary cols=A:Process,B:PQRs,C:Failed rows=0 stale=8
TEST     fixture=pqr_list_01.xlsx call=WritePqrSummary(PqrList, Summary) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND ROWS_USED(Summary) = 6 AND VALUE(Summary![Process]2) = "FCAW" AND VALUE(Summary![PQRs]2) = 1 AND VALUE(Summary![Failed]2) = 1 AND VALUE(Summary![Process]3) = "GTAW" AND VALUE(Summary![PQRs]3) = 4 AND VALUE(Summary![Failed]3) = 1 AND VALUE(Summary![Process]4) = "SAW" AND VALUE(Summary![PQRs]4) = 2 AND VALUE(Summary![Failed]4) = 0 AND VALUE(Summary![Process]5) = "SMAW" AND VALUE(Summary![PQRs]5) = 3 AND VALUE(Summary![Failed]5) = 2 AND VALUE(Summary![Process]6) = "Total" AND VALUE(Summary![PQRs]6) = 10 AND VALUE(Summary![Failed]6) = 4 AND COUNTIF(Summary![Process], "stale") = 0 AND FILL_COUNT(Summary!A1:C9, #FF0000) = 0
TEST     fixture=pqr_list_01.xlsx call=WritePqrSummary(PqrList, Summary) assert=ERROR_COUNT = 0 AND COUNTA(PqrList![Process]) = PRE_COUNTA(PqrList![Process]) AND COUNTA(PqrList![Result]) = PRE_COUNTA(PqrList![Result]) AND COUNTIF(PqrList![Result], "Fail") = PRE_COUNTIF(PqrList![Result], "Fail") AND SUM(Summary![PQRs]) = 20 AND SUM(Summary![Failed]) = 8
TEST     fixture=pqr_list_01.xlsx call=WritePqrSummary(Summary, PqrList) assert=RESULT = -1 AND ERROR_COUNT = 1
