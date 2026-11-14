VCL 0.6
TASK
# ASSUME: the real sheet is "Calibration Tracker"; its row-1 headers include "Instrument ID", "Last Calibrated" and "Next Calibration Due"
# ASSUME: only the Instrument ID cells from row 2 to the last Instrument ID are cleared and coloured; fills elsewhere stay
# ASSUME: a row with a blank Instrument ID is not an instrument: its old highlight is cleared and it is not coloured
# ASSUME: yellow needs real dates in both cells; a due date typed as text, or a missing Last Calibrated, is left uncoloured
# ASSUME: a due date equal to Last Calibrated is not earlier, so it is not coloured
# ASSUME: the number of cells coloured is the macro's return value
GOAL     Highlight the Instrument ID red where the next calibration due date is blank and yellow where it is earlier than the last calibration, and count the cells coloured.
PROC     HighlightCalibrationDateProblems
KIND     Function
MODULE   modCalibration
ARGS     ws As Worksheet
RET      Long
VAR      cId As Long
VAR      cLast As Long
VAR      cNext As Long
VAR      coloured As Long
PRE      COL($ws, "Instrument ID") > 0; COL($ws, "Last Calibrated") > 0; COL($ws, "Next Calibration Due") > 0
MODIFIES $ws fills
OP       LET(cId, COL($ws, "Instrument ID"))
OP       LET(cLast, COL($ws, "Last Calibrated"))
OP       LET(cNext, COL($ws, "Next Calibration Due"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cId) var=r
OP         FORMAT(CELL($ws, r, cId), fill=none)
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cId))
OP         IF IS_EMPTY(CELL($ws, r, cNext))
OP           FORMAT(CELL($ws, r, cId), fill=red)
OP           LET(coloured, coloured + 1)
OP         ELSE
OP           IF IS_DATE(CELL($ws, r, cNext))
OP             IF IS_DATE(CELL($ws, r, cLast))
OP               IF CELL($ws, r, cNext) < CELL($ws, r, cLast)
OP                 FORMAT(CELL($ws, r, cId), fill=yellow)
OP                 LET(coloured, coloured + 1)
OP       RETURN(coloured)
PERF     screen_off
ERR      standard
FIXTURE  cal_dates_01.xlsx sheet=Tracker cols=A:Instrument ID,B:Type,C:Last Calibrated,D:Next Calibration Due rows=10 stale=3 data=A:seq(1001,1)+boundary(1001,1002,1003,1004,"");B:values("Gauge"x4,"Caliper"x3,"Torque wrench"x3);C:dates(2026-03-01,2026-03-31)+boundary(2026-03-10,2026-03-10,2026-03-10,2026-03-10,2026-03-10,"");D:dates(2026-06-01,2026-12-31)+boundary("",2026-03-09,2026-03-10,2026-03-11,"",2026-01-05,"",2026-12-01,2027-01-15,2026-02-01)
FIXTURE  cal_dates_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=cal_dates_01.xlsx call=HighlightCalibrationDateProblems(Tracker) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND FILL_COUNT(Tracker![Instrument ID], #FF0000) = 2 AND FILL_COUNT(Tracker![Instrument ID], #FFFF00) = 2 AND FILL_COUNT(Tracker![Instrument ID]2, #FF0000) = 1 AND FILL_COUNT(Tracker![Instrument ID]8, #FF0000) = 1 AND FILL_COUNT(Tracker![Instrument ID]3, #FFFF00) = 1 AND FILL_COUNT(Tracker![Instrument ID]11, #FFFF00) = 1
TEST     fixture=cal_dates_01.xlsx call=HighlightCalibrationDateProblems(Tracker) assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(Tracker![Instrument ID]) = PRE_COUNTA(Tracker![Instrument ID]) AND COUNTA(Tracker![Next Calibration Due]) = PRE_COUNTA(Tracker![Next Calibration Due]) AND FILL_COUNT(Tracker![Type], #FF0000) = PRE_FILL_COUNT(Tracker![Type], #FF0000)
TEST     fixture=cal_dates_01.xlsx call=HighlightCalibrationDateProblems(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
