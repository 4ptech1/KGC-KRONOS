VCL 0.7
TASK
# ASSUME: The WPS number goes in C6, the revision text next to it in F6, and the WPQR number in C7 of the sheet WPS
# ASSUME: The revision text 'Rev. R0 - first issue' is written as the request gives it, even where the form already says so
GOAL     Write the new WPS number DM-WPS-202, the revision note Rev. R0 - first issue and the WPQR number DM-PQR-202 on the WPS form.
PROC     SetWpsNumbers
KIND     Function
MODULE   modWps
ARGS     ws As Worksheet
RET      Long
VAR      n As Long
PRE      LEFT(CELL($ws, 6, 1), 7) = "Welding"; LEFT(CELL($ws, 7, 1), 4) = "WPQR"
MODIFIES $ws
OP       SET_VALUE(CELL($ws, 6, 3), "DM-WPS-202")
OP       SET_VALUE(CELL($ws, 6, 6), "Rev. R0 - first issue")
OP       SET_VALUE(CELL($ws, 7, 3), "DM-PQR-202")
OP       LET(n, 3)
OP       RETURN(n)
ERR      standard
FIXTURE  wps_01.xlsx template=wps_form
TEST     fixture=wps_01.xlsx call=SetWpsNumbers(WPS) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND VALUE(WPS!C6) = "DM-WPS-202" AND VALUE(WPS!F6) = "Rev. R0 - first issue" AND VALUE(WPS!C7) = "DM-PQR-202"
TEST     fixture=wps_01.xlsx call=SetWpsNumbers(WPS) assert=VALUE(WPS!F7) = "Certificate DM/2026/014" AND VALUE(WPS!C4) = "Demo Fabricators Pvt Ltd" AND VALUE(WPS!H3) = "WPS Revision : R0" AND VALUE(WPS!C37) = "Actual (PQR-201)" AND COUNTA(WPS!A1:AC50) = PRE_COUNTA(WPS!A1:AC50)
