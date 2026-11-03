VCL 0.6
TASK
# ASSUME: WPS Register headers in row 1 include "Process" and "PQR Ref"; each row with a process is one WPS
# ASSUME: the Process Summary sheet already exists and is passed as out; it is cleared and rebuilt on every run
# ASSUME: processes are matched ignoring case and extra spaces at the ends, and the first spelling found is the one listed
# ASSUME: a WPS has no PQR Ref when its PQR Ref cell is blank or only spaces; rows with no process are ignored
# ASSUME: the summary has the headings Process, WPS Count and No PQR Ref, with the processes sorted A to Z
# ASSUME: the out sheet must not hold a PQR Ref heading, so the register itself can never be passed as out and cleared
GOAL     Rebuild the Process Summary sheet with each welding process, how many WPS it has and how many of those have no PQR Ref, sorted A to Z.
PROC     BuildProcessSummary
KIND     Function
MODULE   modWps
ARGS     wps As Worksheet, out As Worksheet
RET      Long
VAR      cProc As Long
VAR      cPqr As Long
VAR      totals As Dictionary
VAR      noPqr As Dictionary
VAR      nTotal As Long
VAR      nNoPqr As Long
VAR      outRow As Long
PRE      COL($wps, "Process") > 0; COL($wps, "PQR Ref") > 0; COL($out, "PQR Ref") = 0
MODIFIES $out
OP       LET(cProc, COL($wps, "Process"))
OP       LET(cPqr, COL($wps, "PQR Ref"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($wps, cProc) var=r
OP         SKIP_IF IS_EMPTY(CELL($wps, r, cProc))
OP         DICT_ADD(totals, TRIM(CELL($wps, r, cProc)), 1, mode=sum)
OP         DICT_ADD(noPqr, TRIM(CELL($wps, r, cProc)), IIF(IS_EMPTY(CELL($wps, r, cPqr)), 1, 0), mode=sum)
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Process", "WPS Count", "No PQR Ref")
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(totals)
OP         DICT_GET(totals, k, into=nTotal)
OP         DICT_GET(noPqr, k, into=nNoPqr)
OP         WRITE_ROW($out, outRow, k, nTotal, nNoPqr)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       RETURN(COUNT(totals))
PERF     screen_off calc_manual
ERR      standard
FIXTURE  wps_01.xlsx sheet=WPS cols=A:WPS No,B:Process,C:PQR Ref,D:Review Date rows=13 data=A:seq(1,1);B:values("SMAW"x3,"GTAW"x2," GTAW"x1,"smaw "x1,"Smaw"x1,"FCAW"x2,"SAW"x1,""x2)+boundary("SMAW","GTAW","smaw ","","FCAW"," GTAW","SAW","","SMAW","FCAW","GTAW","SMAW","Smaw");C:values("PQR-01"x1,"PQR-02"x2,""x4,"PQR-03"x1,"  "x1,"PQR-04"x1,"PQR-05"x1,"PQR-06"x1,"PQR-07"x1)+boundary("PQR-01","PQR-02","","PQR-03","","PQR-02","  ","","PQR-04","PQR-05","PQR-06","","PQR-07");D:reldates(-200,200)
FIXTURE  wps_01.xlsx sheet=Summary cols=A:Process,B:WPS Count,C:No PQR Ref rows=0 stale=3
FIXTURE  wps_01.xlsx sheet=NoProc cols=A:WPS No,B:PQR Ref rows=2 data=A:seq(1,1);B:text(6)
FIXTURE  wps_02.xlsx sheet=WPS cols=A:WPS No,B:Process,C:PQR Ref rows=3 data=A:seq(1,1);B:values(""x3);C:values("PQR-01"x3)
FIXTURE  wps_02.xlsx sheet=Summary cols=A:Process,B:WPS Count,C:No PQR Ref rows=0 stale=2
TEST     fixture=wps_01.xlsx call=BuildProcessSummary(WPS, Summary) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(Summary![Process]) = 5 AND VALUE(Summary![Process]1) = "Process" AND VALUE(Summary![Process]2) = "FCAW" AND VALUE(Summary![WPS Count]2) = 2 AND VALUE(Summary![No PQR Ref]2) = 1 AND VALUE(Summary![Process]3) = "GTAW" AND VALUE(Summary![WPS Count]3) = 3 AND VALUE(Summary![No PQR Ref]3) = 0 AND VALUE(Summary![Process]4) = "SAW" AND VALUE(Summary![WPS Count]4) = 1 AND VALUE(Summary![No PQR Ref]4) = 1 AND VALUE(Summary![Process]5) = "SMAW" AND VALUE(Summary![WPS Count]5) = 5 AND VALUE(Summary![No PQR Ref]5) = 2 AND COUNTIF(Summary![Process], "stale") = 0 AND FILL_COUNT(Summary![Process], #FF0000) = 0
TEST     fixture=wps_01.xlsx call=BuildProcessSummary(WPS, Summary) assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(WPS![Process]) = PRE_COUNTA(WPS![Process]) AND COUNTA(WPS![PQR Ref]) = PRE_COUNTA(WPS![PQR Ref]) AND SUM(Summary![WPS Count]) = 11 AND SUM(Summary![No PQR Ref]) = 4
TEST     fixture=wps_02.xlsx call=BuildProcessSummary(WPS, Summary) assert=RESULT = 0 AND ERROR_COUNT = 0 AND ROWS_USED(Summary) = 1 AND VALUE(Summary![Process]1) = "Process" AND COUNTIF(Summary![Process], "stale") = 0
TEST     fixture=wps_01.xlsx call=BuildProcessSummary(Summary, Summary) assert=RESULT = -1 AND ERROR_COUNT = 1
TEST     fixture=wps_01.xlsx call=BuildProcessSummary(NoProc, Summary) assert=RESULT = -1 AND ERROR_COUNT = 1
TEST     fixture=wps_01.xlsx call=BuildProcessSummary(WPS, WPS) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(WPS![Process]) = PRE_COUNTA(WPS![Process]) AND COUNTA(WPS![PQR Ref]) = PRE_COUNTA(WPS![PQR Ref])
