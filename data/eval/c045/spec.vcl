VCL 0.6
TASK
# ASSUME: Calibration Tracker headers in row 1 include "Next Calibration Due"
# ASSUME: any non-blank cell in that column that is not a real date is marked, including typed numbers and errors
# ASSUME: a cell already marked Date? is counted again on a re-run, so every run returns the same count
GOAL     Write Date? in every Next Calibration Due cell that holds text instead of a real date and return how many cells were marked.
PROC     MarkTextDates
KIND     Function
MODULE   modCalibration
ARGS     ws As Worksheet
RET      Long
VAR      cDue As Long
VAR      marked As Long
PRE      COL($ws, "Next Calibration Due") > 0
MODIFIES $ws["Next Calibration Due"]
OP       LET(cDue, COL($ws, "Next Calibration Due"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cDue) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cDue))
OP         IF NOT IS_DATE(CELL($ws, r, cDue))
OP           SET_VALUE(CELL($ws, r, cDue), "Date?")
OP           LET(marked, marked + 1)
OP       RETURN(marked)
ERR      standard
FIXTURE  calibration_01.xlsx sheet=Calibration cols=A:Instrument ID,B:Next Calibration Due rows=10 data=A:seq(1,1);B:values("TBC"x3,"31/12/2026"x2,"2026-13-45"x1,""x2,"next year"x2)+boundary("TBC","","31/12/2026","")
FIXTURE  calibration_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  calibration_02.xlsx sheet=Calibration cols=A:Instrument ID,B:Next Calibration Due rows=8 data=A:seq(1,1);B:dates(2026-01-01,2026-12-31)+boundary(2026-03-01,"",2026-06-30)
TEST     fixture=calibration_01.xlsx call=MarkTextDates(Calibration) runs=2 assert=RESULT = 8 AND ERROR_COUNT = 0 AND COUNTIF(Calibration![Next Calibration Due], "Date?") = 8
TEST     fixture=calibration_01.xlsx call=MarkTextDates(Calibration) assert=COUNTA(Calibration![Next Calibration Due]) = PRE_COUNTA(Calibration![Next Calibration Due]) AND COUNTA(Calibration![Instrument ID]) = PRE_COUNTA(Calibration![Instrument ID])
TEST     fixture=calibration_01.xlsx call=MarkTextDates(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
TEST     fixture=calibration_02.xlsx call=MarkTextDates(Calibration) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(Calibration![Next Calibration Due], "Date?") = 0 AND COUNTA(Calibration![Next Calibration Due]) = PRE_COUNTA(Calibration![Next Calibration Due])
