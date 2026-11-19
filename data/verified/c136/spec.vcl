VCL 0.6
TASK
# ASSUME: WPS Register headers in row 1 include "WPS No", "Process" and "PQR Ref"; the WPS by Process sheet exists and is passed in, and is rebuilt on every run
# ASSUME: Processes are trimmed and matched ignoring case, and shown in capitals; rows with a blank Process are ignored
# ASSUME: a WPS has no PQR Ref when that cell is empty or only spaces; the result is the number of processes written, not counting Total
GOAL     Rebuild the WPS by Process sheet with each welding process, its number of WPS and how many have no PQR reference, A to Z, ending with a Total row.
PROC     BuildWpsByProcess
KIND     Function
MODULE   modWpsSummary
ARGS     reg As Worksheet, out As Worksheet
RET      Long
VAR      cNo As Long
VAR      cProc As Long
VAR      cPqr As Long
VAR      total As Dictionary
VAR      missing As Dictionary
VAR      nWps As Long
VAR      nMissing As Long
VAR      sumWps As Long
VAR      sumMissing As Long
VAR      outRow As Long
VAR      nProc As Long
PRE      COL($reg, "WPS No") > 0; COL($reg, "Process") > 0; COL($reg, "PQR Ref") > 0
MODIFIES $out
OP       LET(cNo, COL($reg, "WPS No"))
OP       LET(cProc, COL($reg, "Process"))
OP       LET(cPqr, COL($reg, "PQR Ref"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Process", "WPS Count", "No PQR Ref")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($reg, cNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($reg, r, cProc))
OP         DICT_ADD(total, UPPER(TRIM(CELL($reg, r, cProc))), 1, mode=sum)
OP         DICT_ADD(missing, UPPER(TRIM(CELL($reg, r, cProc))), IIF(IS_EMPTY(CELL($reg, r, cPqr)), 1, 0), mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(total)
OP         DICT_GET(total, k, into=nWps)
OP         DICT_GET(missing, k, into=nMissing)
OP         WRITE_ROW($out, outRow, k, nWps, nMissing)
OP         LET(sumWps, sumWps + nWps)
OP         LET(sumMissing, sumMissing + nMissing)
OP         LET(outRow, outRow + 1)
OP       LET(nProc, outRow - 2)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       WRITE_ROW($out, outRow, "Total", sumWps, sumMissing)
OP       RETURN(nProc)
PERF     screen_off
ERR      standard
FIXTURE  wps_01.xlsx sheet=Reg cols=A:WPS No,B:Process,C:PQR Ref rows=12 data=A:seq(1,1);B:values("gtaw"x1,"GTAW"x2,"SMAW"x2," smaw "x1,"Gtaw"x1,"FCAW"x1,"SAW"x2,"smaw"x1,""x1)+boundary("gtaw","GTAW","SMAW"," smaw ","Gtaw","FCAW","SAW","smaw","GTAW","SAW","","SMAW");C:values("PQR-1"x3,"PQR-2"x1,"PQR-3"x2,"PQR-4"x1,""x4," "x1)+boundary("PQR-1","","PQR-2"," ","PQR-1","","PQR-3","","PQR-1","PQR-3","","PQR-4")
FIXTURE  wps_01.xlsx sheet=Out cols=A:Process,B:WPS Count,C:No PQR Ref rows=0 stale=3
TEST     fixture=wps_01.xlsx call=BuildWpsByProcess(Reg, Out) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(Out![Process]) = 6 AND COUNTIF(Out![Process], "stale") = 0 AND VALUE(Out![Process]2) = "FCAW" AND VALUE(Out![WPS Count]2) = 1 AND VALUE(Out![No PQR Ref]2) = 1 AND VALUE(Out![Process]3) = "GTAW" AND VALUE(Out![WPS Count]3) = 4 AND VALUE(Out![No PQR Ref]3) = 1
TEST     fixture=wps_01.xlsx call=BuildWpsByProcess(Reg, Out) assert=COUNTA(Reg![WPS No]) = PRE_COUNTA(Reg![WPS No]) AND COUNTA(Reg![PQR Ref]) = PRE_COUNTA(Reg![PQR Ref]) AND VALUE(Out![Process]4) = "SAW" AND VALUE(Out![WPS Count]4) = 2 AND VALUE(Out![No PQR Ref]4) = 0 AND VALUE(Out![Process]5) = "SMAW" AND VALUE(Out![WPS Count]5) = 4 AND VALUE(Out![No PQR Ref]5) = 2 AND VALUE(Out![Process]6) = "Total" AND VALUE(Out![WPS Count]6) = 11 AND VALUE(Out![No PQR Ref]6) = 4 AND ROWS_USED(Out) = 6
TEST     fixture=wps_01.xlsx call=BuildWpsByProcess(Out, Reg) assert=RESULT = -1 AND ERROR_COUNT = 1
