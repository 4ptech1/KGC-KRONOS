VCL 0.7
TASK
# ASSUME: The WPS form's remarks are in B46 and B47.
GOAL     Write two remarks on the practice WPS form, worded by the AI for an engineer to check.
PROC     WriteWpsRemarks
KIND     Function
MODULE   modWps
ARGS     ws As Worksheet
RET      Long
PRE      SAME(CELL($ws, 46, 1), "Remarks :")
MODIFIES $ws
OP       SET_VALUE(RANGE($ws, "B46"), AI_TEXT("1. Grind the root back to sound metal before side B."))
OP       SET_VALUE(RANGE($ws, "B47"), "2. " & AI_TEXT("Keep the interpass temperature at or below 250 °C."))
OP       RETURN(2)
ERR      standard
FIXTURE  wps_02.xlsx template=wps_form
TEST     fixture=wps_02.xlsx call=WriteWpsRemarks(WPS) runs=2 assert=RESULT = 2 AND VALUE(WPS!B46) = "1. Grind the root back to sound metal before side B." AND VALUE(WPS!B47) = "2. Keep the interpass temperature at or below 250 °C." AND ERROR_COUNT = 0
TEST     fixture=wps_02.xlsx call=WriteWpsRemarks(WPS) assert=COUNTA(WPS!A1:I50) = PRE_COUNTA(WPS!A1:I50) AND VALUE(WPS!A46) = "Remarks :"
