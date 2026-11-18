VCL 0.6
TASK
# ASSUME: Calibration Tracker headers in row 1 include "Instrument ID", "Range Max", "Reading Max" and "Remark"
# ASSUME: a value counts as a number when IS_NUMBER accepts it, so numbers typed as text are compared as numbers
# ASSUME: a reading equal to the range maximum is not over range
# ASSUME: every row up to the last Instrument ID is checked, and the Remark of a row that is neither is cleared
GOAL     Write Over range or Check in Remark for each instrument whose reading is above its range or cannot be compared, clear the other remarks and return the Over range count.
PROC     MarkOverRange
KIND     Function
MODULE   modCalibration
ARGS     ws As Worksheet
RET      Long
VAR      cId As Long
VAR      cRange As Long
VAR      cReading As Long
VAR      cRemark As Long
VAR      over As Long
PRE      COL($ws, "Instrument ID") > 0; COL($ws, "Range Max") > 0; COL($ws, "Reading Max") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Remark"]
OP       LET(cId, COL($ws, "Instrument ID"))
OP       LET(cRange, COL($ws, "Range Max"))
OP       LET(cReading, COL($ws, "Reading Max"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cId) var=r
OP         IF IS_NUMBER(CELL($ws, r, cRange))
OP           IF IS_NUMBER(CELL($ws, r, cReading))
OP             IF CELL($ws, r, cReading) > CELL($ws, r, cRange)
OP               SET_VALUE(CELL($ws, r, cRemark), "Over range")
OP               LET(over, over + 1)
OP             ELSE
OP               CLEAR(CELL($ws, r, cRemark), what=contents)
OP           ELSE
OP             SET_VALUE(CELL($ws, r, cRemark), "Check")
OP         ELSE
OP           SET_VALUE(CELL($ws, r, cRemark), "Check")
OP       RETURN(over)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  cal_01.xlsx sheet=Cal cols=A:Instrument ID,B:Type,C:Range Max,D:Reading Max,E:Remark rows=10 data=A:seq(1,1);B:values("Gauge"x5,"Meter"x5);C:values("100"x6,""x1,"abc"x1,"9"x1,"20"x1)+boundary("100","100","100","","100","abc","100","9","100","20");D:values("120"x1,"100"x1,"80"x1,"50"x2,""x1,"n/a"x1,"10"x1,"99"x1,"5"x1)+boundary("120","100","80","50","","50","n/a","10","99","5");E:values("Over range"x1,"Check"x1,"old"x1,""x7)+boundary("","Over range","Check","","","","","","old","")
FIXTURE  cal_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=cal_01.xlsx call=MarkOverRange(Cal) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND COUNTIF(Cal![Remark], "Over range") = 2 AND COUNTIF(Cal![Remark], "Check") = 4 AND COUNTIF(Cal![Remark], "old") = 0
TEST     fixture=cal_01.xlsx call=MarkOverRange(Cal) assert=COUNTA(Cal![Remark]) = 7 AND VALUE(Cal![Remark]2) = "Over range" AND VALUE(Cal![Remark]3) = "" AND VALUE(Cal![Remark]4) = "" AND VALUE(Cal![Remark]5) = "Check" AND VALUE(Cal![Remark]6) = "Check" AND VALUE(Cal![Remark]7) = "Check" AND VALUE(Cal![Remark]8) = "Check" AND VALUE(Cal![Remark]9) = "Over range" AND VALUE(Cal![Remark]10) = "" AND VALUE(Cal![Remark]11) = ""
TEST     fixture=cal_01.xlsx call=MarkOverRange(Cal) assert=COUNTA(Cal![Instrument ID]) = PRE_COUNTA(Cal![Instrument ID]) AND COUNTA(Cal![Range Max]) = PRE_COUNTA(Cal![Range Max]) AND COUNTA(Cal![Reading Max]) = PRE_COUNTA(Cal![Reading Max])
TEST     fixture=cal_01.xlsx call=MarkOverRange(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
