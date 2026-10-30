VCL 0.6
TASK
# ASSUME: the sheet is the Calibration Tracker with the headers Instrument ID, Last Calibrated, Interval Days and Next Due in row 1, in that order (columns A to D)
# ASSUME: a data row is every row from 2 down to the last row that has an Instrument ID
# ASSUME: Next Due is Last Calibrated plus Interval Days, and stays blank when either of them is not a number
# ASSUME: Next Due cells are already formatted as dates; no format is changed
GOAL     Put a formula in the Next Due column of every data row of the Calibration Tracker that adds Interval Days to Last Calibrated.
PROC     SetNextDueFormulas
KIND     Function
MODULE   modCalibration
ARGS     ws As Worksheet
RET      Long
VAR      cId As Long
VAR      cDue As Long
VAR      n As Long
PRE      COL($ws, "Instrument ID") > 0; COL($ws, "Last Calibrated") = 2; COL($ws, "Interval Days") = 3; COL($ws, "Next Due") = 4
MODIFIES $ws["Next Due"]
OP       LET(cId, COL($ws, "Instrument ID"))
OP       LET(cDue, COL($ws, "Next Due"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cId) var=r
OP         SET_FORMULA(CELL($ws, r, cDue), "=IF(COUNT(RC2:RC3)=2,RC2+RC3,"""")", style=r1c1)
OP         LET(n, n + 1)
OP       RETURN(n)
ERR      standard
FIXTURE  cal_01.xlsx sheet=CalTracker cols=A:Instrument ID,B:Last Calibrated,C:Interval Days,D:Next Due rows=8 data=A:text(6);B:dates(2026-01-01,2026-06-30)+blank(1);C:int(30,365);D:values("old"x8)
FIXTURE  cal_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  cal_02.xlsx sheet=CalTracker cols=A:Instrument ID,B:Interval Days,C:Last Calibrated,D:Next Due rows=3 data=A:text(6);B:int(30,365);C:dates(2026-01-01,2026-06-30);D:values("old"x3)
TEST     fixture=cal_01.xlsx call=SetNextDueFormulas(CalTracker) runs=2 assert=RESULT = 8 AND ERROR_COUNT = 0 AND FORMULA(CalTracker![Next Due]2) = "=IF(COUNT(RC2:RC3)=2,RC2+RC3,"""")" AND FORMULA(CalTracker![Next Due]5) = "=IF(COUNT(RC2:RC3)=2,RC2+RC3,"""")" AND FORMULA(CalTracker![Next Due]9) = "=IF(COUNT(RC2:RC3)=2,RC2+RC3,"""")" AND FORMULA(CalTracker![Next Due]10) = "" AND FORMULA(CalTracker![Next Due]1) = ""
TEST     fixture=cal_01.xlsx call=SetNextDueFormulas(CalTracker) assert=RESULT = 8 AND FORMULA(CalTracker![Instrument ID]2) = "" AND FORMULA(CalTracker![Last Calibrated]2) = "" AND FORMULA(CalTracker![Interval Days]2) = "" AND COUNTA(CalTracker![Instrument ID]) = PRE_COUNTA(CalTracker![Instrument ID]) AND COUNTA(CalTracker![Last Calibrated]) = PRE_COUNTA(CalTracker![Last Calibrated]) AND SUM(CalTracker![Interval Days]) = PRE_SUM(CalTracker![Interval Days])
TEST     fixture=cal_01.xlsx call=SetNextDueFormulas(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
TEST     fixture=cal_02.xlsx call=SetNextDueFormulas(CalTracker) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTIF(CalTracker![Next Due], "old") = PRE_COUNTIF(CalTracker![Next Due], "old")
