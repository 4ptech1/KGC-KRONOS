VCL 0.7
TASK
# ASSUME: Both WPS use the same form; the old WPS's heat-input block holds the PQR's runs, and its T column says OK for a run inside the limit.
GOAL     Fill a new WPS form's PQR number and run table from an old WPS, and note how many old runs were inside the heat-input limit.
PROC     FillWpsFromOld
KIND     Function
MODULE   modWps
ARGS     old As Workbook, ws As Worksheet
RET      Long
VAR      n As Long
PRE      SHEET_EXISTS($old, "WPS"); SAME(CELL($ws, 7, 1), "WPQR No. :")
MODIFIES $ws
OP       SET_VALUE(RANGE($ws, "C7"), CELL(SHEET($old, "WPS"), 7, 3))
OP       LOOP_ROWS dir=down from=29 to=31 var=r
OP         LOOP_COLS dir=right from=12 to=17 var=c
OP           SET_VALUE(CELL($ws, r, c), CELL(SHEET($old, "WPS"), r, c))
OP         IF SAME(CELL(SHEET($old, "WPS"), r, 20), "OK")
OP           LET(n, n + 1)
OP       SET_VALUE(RANGE($ws, "B47"), "2. Runs inside the heat-input limit in the old WPS: " & n)
OP       RETURN(n)
ERR      standard
FIXTURE  wps_old_01.xlsx template=wps_form_b
FIXTURE  wps_new_01.xlsx template=wps_form
TEST     fixture=wps_new_01.xlsx read=wps_old_01.xlsx call=FillWpsFromOld(wps_old_01.xlsx, WPS) runs=2 assert=RESULT = 3 AND VALUE(WPS!C7) = "DM-PQR-105" AND VALUE(WPS!L29) = 80 AND VALUE(WPS!Q31) = 120 AND VALUE(WPS!B47) = "2. Runs inside the heat-input limit in the old WPS: 3" AND ERROR_COUNT = 0
TEST     fixture=wps_new_01.xlsx read=wps_old_01.xlsx call=FillWpsFromOld(wps_old_01.xlsx, WPS) assert=COUNTA(WPS!A1:I50) = PRE_COUNTA(WPS!A1:I50) AND VALUE(WPS!C6) = "DM-WPS-201" AND FORMULA(WPS!S31) = "=$N$25*O31*M31*60/(P31*1000)"
