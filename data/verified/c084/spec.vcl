VCL 0.6
TASK
# ASSUME: WPS Register headers in row 1 include "Process" and "PQR Ref"
# ASSUME: trimmed means the spaces at both ends are removed; spaces inside a Process are kept
# ASSUME: a Process cell that holds only spaces becomes empty and counts as one changed cell
# ASSUME: a Process cell that is already trimmed and upper case, or holds a number, is not written and not counted
# ASSUME: only spaces also means non-breaking spaces, tabs and line breaks; a PQR Ref that holds any other text is left exactly as typed
GOAL     Trim and upper-case every Process cell of the WPS Register, empty every PQR Ref cell that holds only spaces and return how many cells were changed.
PROC     CleanWpsProcessAndPqrRef
KIND     Function
MODULE   modWps
ARGS     ws As Worksheet
RET      Long
VAR      cProc As Long
VAR      cPqr As Long
VAR      n As Long
PRE      COL($ws, "Process") > 0; COL($ws, "PQR Ref") > 0
MODIFIES $ws["Process", "PQR Ref"]
OP       LET(cProc, COL($ws, "Process"))
OP       LET(cPqr, COL($ws, "PQR Ref"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cProc) var=r
OP         IF CELL($ws, r, cProc) <> UPPER(TRIM(CELL($ws, r, cProc)))
OP           SET_VALUE(CELL($ws, r, cProc), UPPER(TRIM(CELL($ws, r, cProc))))
OP           LET(n, n + 1)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cPqr) var=q
OP         IF IS_EMPTY(CELL($ws, q, cPqr)) AND LEN(CELL($ws, q, cPqr)) > 0
OP           SET_VALUE(CELL($ws, q, cPqr), "")
OP           LET(n, n + 1)
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  wps_clean_01.xlsx sheet=Wps cols=A:WPS No,B:Process,C:PQR Ref rows=10 data=A:seq(101,1);B:values("smaw"x1," SMAW"x1,"Gtaw "x1,"GTAW"x1,"SAW"x1,"  fcaw  "x1,""x1,"   "x1,"GMAW"x1,"gtaw"x1)+boundary("smaw"," SMAW","Gtaw ","GTAW","SAW","  fcaw  ","","   ","GMAW","gtaw");C:values(" "x1,"PQR-01"x1,""x2,"   "x1,"PQR-02 "x1,"PQR-03"x1,"pqr-04"x1,"PQR-05"x1,"  "x1)+boundary(" ","PQR-01","","   ","PQR-02 ","PQR-03","","pqr-04","PQR-05","  ")
FIXTURE  wps_clean_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=wps_clean_01.xlsx call=CleanWpsProcessAndPqrRef(Wps) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE(Wps![Process]2) = "SMAW" AND VALUE(Wps![Process]3) = "SMAW" AND VALUE(Wps![Process]4) = "GTAW" AND VALUE(Wps![Process]7) = "FCAW" AND VALUE(Wps![Process]9) = "" AND VALUE(Wps![Process]11) = "GTAW" AND COUNTIF(Wps![Process], "SMAW") = 2 AND COUNTIF(Wps![Process], "GTAW") = 3 AND VALUE(Wps![PQR Ref]2) = "" AND VALUE(Wps![PQR Ref]5) = "" AND VALUE(Wps![PQR Ref]11) = ""
TEST     fixture=wps_clean_01.xlsx call=CleanWpsProcessAndPqrRef(Wps) assert=RESULT = 9 AND ERROR_COUNT = 0 AND COUNTA(Wps![WPS No]) = PRE_COUNTA(Wps![WPS No]) AND SUM(Wps![WPS No]) = PRE_SUM(Wps![WPS No]) AND VALUE(Wps![Process]5) = "GTAW" AND VALUE(Wps![Process]6) = "SAW" AND VALUE(Wps![Process]10) = "GMAW" AND VALUE(Wps![PQR Ref]3) = "PQR-01" AND VALUE(Wps![PQR Ref]6) = "PQR-02 " AND VALUE(Wps![PQR Ref]9) = "pqr-04" AND COUNTIF(Wps![PQR Ref], "PQR-0*") = 5
TEST     fixture=wps_clean_01.xlsx call=CleanWpsProcessAndPqrRef(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
