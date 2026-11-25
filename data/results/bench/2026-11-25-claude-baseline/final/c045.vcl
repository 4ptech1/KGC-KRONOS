VCL 0.7
TASK
# ASSUME: Calibration Tracker headers in row 1 include "Next Calibration Due"
# ASSUME: any non-blank cell in that column that is not a real date is marked Date? (text, including text that looks like a date); real dates and blank cells are left alone
# ASSUME: a cell already holding Date? is marked and counted again on a re-run
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
OP         SKIP_IF IS_DATE(CELL($ws, r, cDue))
OP         SET_VALUE(CELL($ws, r, cDue), "Date?")
OP         LET(marked, marked + 1)
OP       RETURN(marked)
PERF     screen_off
ERR      standard
FIXTURE  cal_01.xlsx sheet=Cal cols=A:Instrument ID,B:Next Calibration Due rows=6 data=A:seq(1,1);B:values("31/12/2026"x2,"TBC"x2,"n/a"x1,""x1)+boundary("31/12/2026","TBC","","n/a","TBC","31/12/2026")
FIXTURE  cal_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  cal_02.xlsx sheet=Cal cols=A:Instrument ID,B:Next Calibration Due rows=6 data=A:seq(1,1);B:dates(2026-01-01,2027-12-31)+boundary("")
FIXTURE  cal_02.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=cal_01.xlsx call=MarkTextDates(Cal) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTIF(Cal![Next Calibration Due], "Date?") = 5
TEST     fixture=cal_01.xlsx call=MarkTextDates(Cal) assert=COUNTA(Cal![Instrument ID]) = PRE_COUNTA(Cal![Instrument ID]) AND COUNTA(Cal![Next Calibration Due]) = PRE_COUNTA(Cal![Next Calibration Due]) AND COUNTIF(Cal![Next Calibration Due], "TBC") = 0
TEST     fixture=cal_02.xlsx call=MarkTextDates(Cal) assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTA(Cal![Next Calibration Due]) = PRE_COUNTA(Cal![Next Calibration Due])
TEST     fixture=cal_01.xlsx call=MarkTextDates(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
