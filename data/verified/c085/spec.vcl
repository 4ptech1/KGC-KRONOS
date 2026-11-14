VCL 0.6
TASK
# ASSUME: the real sheet is "Calibration Tracker"; its row-1 headers include "Status"
# ASSUME: a status is matched ignoring case, spaces at the ends and extra spaces between its words
# ASSUME: only a whole-cell match is rewritten; "in calibration" or "not ok" is another status and stays
# ASSUME: a cell that already reads exactly Calibrated or Out of Service is not rewritten and not counted
# ASSUME: the number of cells changed is the macro's return value
GOAL     Rewrite every Status that means calibrated as exactly Calibrated and every one that means out of service as exactly Out of Service, and count the cells changed.
PROC     StandardiseCalibrationStatus
KIND     Function
MODULE   modCalibration
ARGS     ws As Worksheet
RET      Long
VAR      cStatus As Long
VAR      typed As String
VAR      want As String
VAR      changed As Long
PRE      COL($ws, "Status") > 0
MODIFIES $ws["Status"]
OP       LET(cStatus, COL($ws, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cStatus) var=r
OP         LET(typed, CLEAN_TEXT(CELL($ws, r, cStatus)))
OP         LET(want, "")
OP         IF SAME(typed, "ok") OR SAME(typed, "in cal") OR SAME(typed, "calibrated") THEN LET(want, "Calibrated")
OP         IF SAME(typed, "oos") OR SAME(typed, "out of service") THEN LET(want, "Out of Service")
OP         SKIP_IF want = ""
OP         IF CELL($ws, r, cStatus) <> want
OP           SET_VALUE(CELL($ws, r, cStatus), want)
OP           LET(changed, changed + 1)
OP       RETURN(changed)
PERF     screen_off
ERR      standard
FIXTURE  cal_status_01.xlsx sheet=Tracker cols=A:Instrument ID,B:Type,C:Location,D:Status rows=14 data=A:seq(2001,1);B:values("Gauge"x6,"Caliper"x4,"Torque wrench"x4);C:values("Bay 1"x5,"Bay 2"x5,"Stores"x4);D:values(" OK "x1,"Quarantined"x1,"In  Cal"x1,""x1,"CALIBRATED"x1,"Calibrated"x1,"oos "x1,"Not ok"x1," Out  of  SERVICE"x1,"Out of Service"x1,"in calibration"x1,"OOS"x1,"ok"x1,"calibrated"x1)+boundary(" OK ","Quarantined","In  Cal","","CALIBRATED","Calibrated","oos ","Not ok"," Out  of  SERVICE","Out of Service","in calibration","OOS","ok","calibrated")
FIXTURE  cal_status_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=cal_status_01.xlsx call=StandardiseCalibrationStatus(Tracker) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(Tracker![Status], "Calibrated") = 6 AND COUNTIF(Tracker![Status], "Out of Service") = 4 AND VALUE(Tracker![Status]2) = "Calibrated" AND VALUE(Tracker![Status]4) = "Calibrated" AND VALUE(Tracker![Status]6) = "Calibrated" AND VALUE(Tracker![Status]8) = "Out of Service" AND VALUE(Tracker![Status]10) = "Out of Service" AND VALUE(Tracker![Status]13) = "Out of Service" AND VALUE(Tracker![Status]14) = "Calibrated" AND VALUE(Tracker![Status]15) = "Calibrated"
TEST     fixture=cal_status_01.xlsx call=StandardiseCalibrationStatus(Tracker) assert=RESULT = 8 AND ERROR_COUNT = 0 AND COUNTA(Tracker![Status]) = PRE_COUNTA(Tracker![Status]) AND COUNTA(Tracker![Instrument ID]) = PRE_COUNTA(Tracker![Instrument ID]) AND VALUE(Tracker![Status]3) = "Quarantined" AND VALUE(Tracker![Status]9) = "Not ok" AND VALUE(Tracker![Status]12) = "in calibration" AND COUNTIF(Tracker![Status], "Calibrated") = PRE_COUNTIF(Tracker![Status], "Calibrated") + 3
TEST     fixture=cal_status_01.xlsx call=StandardiseCalibrationStatus(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
