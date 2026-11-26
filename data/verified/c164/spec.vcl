VCL 0.7
TASK
# ASSUME: The block is the PQR columns V to AA of rows 29 to 31 on the sheet WPS: PQR I min, I max, U min, U max, v and HI recorded
# ASSUME: Only contents are cleared; the formulas in AB and AC and every other cell stay
GOAL     Clear the recorded PQR values of runs 1 to 3 (PQR I, U, v and HI recorded) in the heat-input check block of the WPS.
PROC     ClearPqrRunValues
KIND     Function
MODULE   modWps
ARGS     ws As Worksheet
RET      Long
VAR      n As Long
PRE      LEFT(CELL($ws, 28, 22), 3) = "PQR"; LEFT(CELL($ws, 28, 27), 3) = "PQR"; LEFT(CELL($ws, 28, 28), 3) = "PQR"
MODIFIES $ws
OP       LOOP_ROWS dir=down from=29 to=31 var=r
OP         CLEAR(ROW_RANGE($ws, r, 22, 27), what=contents)
OP         LET(n, n + 1)
OP       RETURN(n)
ERR      standard
FIXTURE  wps_01.xlsx template=wps_form
TEST     fixture=wps_01.xlsx call=ClearPqrRunValues(WPS) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTA(WPS!V28:AA31) = 6 AND COUNTA(WPS!V29:AA29) = 0 AND COUNTA(WPS!V30:AA30) = 0 AND COUNTA(WPS!V31:AA31) = 0
TEST     fixture=wps_01.xlsx call=ClearPqrRunValues(WPS) assert=VALUE(WPS!V28) = "PQR I min" AND VALUE(WPS!AA28) = "PQR HI recorded" AND VALUE(WPS!L29) = 90 AND VALUE(WPS!Q31) = 130 AND VALUE(WPS!N26) = 1.6 AND COUNTA(WPS!A1:AC50) = PRE_COUNTA(WPS!A1:AC50) - 18
