VCL 0.6
TASK
# ASSUME: WPS Register headers in row 1 include "WPS No", "Thickness Min", "Thickness Max", "PQR No" and "Remark"
# ASSUME: the Remark column already exists in row 1; the macro never adds it
# ASSUME: every row from 2 down to the last WPS No is checked
# ASSUME: a PQR No holding only spaces counts as blank
# ASSUME: a thickness that is not a number cannot be compared, so on its own it does not give Check
# ASSUME: the real sheet is "WPS Register"; the TESTs use WpsReg
GOAL     Write Check in the Remark of every WPS row with a blank PQR No or a Thickness Min above its Thickness Max, clear Remark on the other rows and tell how many have Check.
PROC     MarkWpsChecks
KIND     Function
MODULE   modWps
ARGS     ws As Worksheet
RET      Long
VAR      cWps As Long
VAR      cMin As Long
VAR      cMax As Long
VAR      cPqr As Long
VAR      cRemark As Long
VAR      n As Long
PRE      COL($ws, "WPS No") > 0; COL($ws, "Thickness Min") > 0; COL($ws, "Thickness Max") > 0; COL($ws, "PQR No") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Remark"]
OP       LET(cWps, COL($ws, "WPS No"))
OP       LET(cMin, COL($ws, "Thickness Min"))
OP       LET(cMax, COL($ws, "Thickness Max"))
OP       LET(cPqr, COL($ws, "PQR No"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cWps) var=r
OP         CLEAR(CELL($ws, r, cRemark), what=contents)
OP         IF IS_EMPTY(CELL($ws, r, cPqr))
OP           SET_VALUE(CELL($ws, r, cRemark), "Check")
OP           LET(n, n + 1)
OP         ELSE
OP           IF IS_NUMBER(CELL($ws, r, cMin))
OP             IF IS_NUMBER(CELL($ws, r, cMax))
OP               IF CELL($ws, r, cMin) > CELL($ws, r, cMax)
OP                 SET_VALUE(CELL($ws, r, cRemark), "Check")
OP                 LET(n, n + 1)
OP       MSG("WPS rows with Check: " & n)
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  wps_01.xlsx sheet=WpsReg cols=A:WPS No,B:Process,C:Thickness Min,D:Thickness Max,E:PQR No,F:Remark rows=12 data=A:seq(101,1);B:values("GTAW"x5,"SMAW"x4,"SAW"x3);C:int(3,20)+boundary(6,8,9,6,10,5,"",4,12,3,7,2);D:int(5,25)+boundary(12,8,8,12,5,10,10,"",20,25,6,4);E:values("PQR-01"x1,"PQR-02"x1,"PQR-03"x1," "x1,""x1,"PQR-04"x1,"PQR-05"x1,"PQR-06"x1,"PQR-07"x1,"PQR-08"x1,"PQR-09"x1,"PQR-10"x1)+boundary("PQR-01","PQR-02","PQR-03"," ","","PQR-04","PQR-05","PQR-06","PQR-07","PQR-08","PQR-09","PQR-10");F:values("Check"x5,"Old"x2,""x5)+boundary("Check","Check","","Check","","Check","Check","Old","Old","","","")
FIXTURE  wps_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=wps_01.xlsx call=MarkWpsChecks(WpsReg) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(WpsReg![Remark], "Check") = 4 AND COUNTIF(WpsReg![Remark], "Old") = 0 AND COUNTA(WpsReg![Remark]) = 5
TEST     fixture=wps_01.xlsx call=MarkWpsChecks(WpsReg) assert=VALUE(WpsReg![Remark]4) = "Check" AND VALUE(WpsReg![Remark]5) = "Check" AND VALUE(WpsReg![Remark]6) = "Check" AND VALUE(WpsReg![Remark]12) = "Check" AND COUNTIFS(WpsReg![PQR No], "PQR-02", WpsReg![Remark], "Check") = 0 AND COUNTA(WpsReg![WPS No]) = PRE_COUNTA(WpsReg![WPS No]) AND COUNTA(WpsReg![PQR No]) = PRE_COUNTA(WpsReg![PQR No]) AND SUM(WpsReg![Thickness Min]) = PRE_SUM(WpsReg![Thickness Min])
TEST     fixture=wps_01.xlsx call=MarkWpsChecks(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
