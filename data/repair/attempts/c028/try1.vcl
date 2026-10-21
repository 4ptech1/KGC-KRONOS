VCL 0.6
TASK
# ASSUME: the Calibration Tracker sheet has "Instrument ID" in A1 and one instrument per row below it
# ASSUME: the new sheet is named exactly Calibration Totals and goes after the last sheet
# ASSUME: A1 counts every non-empty cell in the Instrument ID column below the heading
# ASSUME: when Calibration Totals is already there the run changes nothing and returns 0
GOAL     Add a Calibration Totals sheet with a formula in A1 that counts the instruments in the Calibration Tracker, only when that sheet is not there yet.
PROC     AddCalibrationTotals
KIND     Function
MODULE   modCalibration
ARGS     book As Workbook, tracker As Worksheet
RET      Long
VAR      totals As Worksheet
PRE      COL($tracker, "Instrument ID") = 1
MODIFIES $book sheets
OP       EXIT_IF SHEET_EXISTS($book, "Calibration Totals")
OP       ADD_SHEET($book, name="Calibration Totals", into=totals)
OP       SET_FORMULA(CELL(totals, 1, 1), "=COUNTA('Calibration Tracker'!A:A)-1", style=a1)
OP       RETURN(1)
ERR      standard
FIXTURE  cal_01.xlsx sheet=CalTracker cols=A:Instrument ID,B:Next Due rows=6 data=A:seq(1,1);B:dates(2026-01-01,2026-12-31)
FIXTURE  cal_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=cal_01.xlsx call=AddCalibrationTotals(cal_01.xlsx, CalTracker) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND SHEET_EXISTS("Calibration Totals")
TEST     fixture=cal_01.xlsx call=AddCalibrationTotals(cal_01.xlsx, CalTracker) assert=RESULT = 1 AND ERROR_COUNT = 0 AND FORMULA('Calibration Totals'!A1) = "=COUNTA('Calibration Tracker'!A:A)-1" AND COUNTA(CalTracker![Instrument ID]) = PRE_COUNTA(CalTracker![Instrument ID])
TEST     fixture=cal_01.xlsx call=AddCalibrationTotals(cal_01.xlsx, Notes) assert=RESULT = -1 AND ERROR_COUNT = 1 AND SHEET_EXISTS("Calibration Totals") = 0
