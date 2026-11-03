VCL 0.6
TASK
# ASSUME: WPS Register headers in row 1 include "WPS No", "PQR Ref" and "Remark"
# ASSUME: a row with a blank WPS No is not a WPS: it is skipped and its Remark is left alone
# ASSUME: "starts with PQR-" ignores case and stray spaces at the start, so " pqr-12" is fine; a blank PQR Ref is marked
# ASSUME: the count is every WPS marked Check by this run, whether or not it already said Check, so a re-run gives the same count
GOAL     Write Check in Remark for every WPS whose PQR Ref does not start with PQR-, clear Remark on the others, and tell how many were marked.
PROC     FlagWpsWithoutPqr
KIND     Function
MODULE   modWps
ARGS     ws As Worksheet
RET      Long
VAR      cWps As Long
VAR      cPqr As Long
VAR      cRemark As Long
VAR      marked As Long
PRE      COL($ws, "WPS No") > 0; COL($ws, "PQR Ref") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Remark"]
OP       LET(cWps, COL($ws, "WPS No"))
OP       LET(cPqr, COL($ws, "PQR Ref"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cWps) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cWps))
OP         IF SAME(LEFT(TRIM(CELL($ws, r, cPqr)), 4), "PQR-")
OP           CLEAR(CELL($ws, r, cRemark), what=contents)
OP         ELSE
OP           SET_VALUE(CELL($ws, r, cRemark), "Check")
OP           LET(marked, marked + 1)
OP       MSG("WPS marked Check: " & marked)
OP       RETURN(marked)
ERR      standard
FIXTURE  wps_01.xlsx sheet=Wps cols=A:WPS No,B:Process,C:PQR Ref,D:Review Date,E:Remark rows=12 data=A:seq(1,1)+boundary(1,2,3,4,5,6,7,8,9,10,"",12);B:values("SMAW"x4,"GTAW"x4,"FCAW"x4);C:values("PQR-101"x1,"PQR-102"x1,"pqr-103"x1," PQR-104"x1,"PQR-"x1,"PQR"x1,"PQR_105"x1,"XPQR-106"x1,""x1,"   "x1,"ABC-1"x1,"N/A"x1)+boundary("PQR-101","PQR-102","pqr-103"," PQR-104","PQR-","PQR","PQR_105","XPQR-106","","   ","ABC-1","N/A");D:dates(2026-01-01,2026-09-30);E:values("Check"x3,"ok"x1,"old"x1,"keep"x1,""x6)+boundary("Check","","ok","Check","","","old","","","Check","keep","")
FIXTURE  wps_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=wps_01.xlsx call=FlagWpsWithoutPqr(Wps) runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND COUNTIF(Wps![Remark], "Check") = 6 AND COUNTA(Wps![Remark]) = 8
TEST     fixture=wps_01.xlsx call=FlagWpsWithoutPqr(Wps) assert=RESULT = 6 AND COUNTA(Wps![WPS No]) = PRE_COUNTA(Wps![WPS No]) AND COUNTA(Wps![PQR Ref]) = PRE_COUNTA(Wps![PQR Ref]) AND COUNTA(Wps![Process]) = PRE_COUNTA(Wps![Process]) AND COUNTA(Wps![Review Date]) = PRE_COUNTA(Wps![Review Date])
TEST     fixture=wps_01.xlsx call=FlagWpsWithoutPqr(Wps) assert=VALUE(Wps![Remark]2) = "" AND VALUE(Wps![Remark]4) = "" AND VALUE(Wps![Remark]5) = "" AND VALUE(Wps![Remark]6) = "" AND VALUE(Wps![Remark]7) = "Check" AND VALUE(Wps![Remark]8) = "Check" AND VALUE(Wps![Remark]9) = "Check" AND VALUE(Wps![Remark]10) = "Check" AND VALUE(Wps![Remark]11) = "Check" AND VALUE(Wps![Remark]12) = "keep" AND VALUE(Wps![Remark]13) = "Check"
TEST     fixture=wps_01.xlsx call=FlagWpsWithoutPqr(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
