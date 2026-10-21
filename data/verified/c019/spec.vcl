VCL 0.6
TASK
# ASSUME: Calibration Tracker headers in row 1 include "Next Calibration Due"
# ASSUME: the real sheets are "Calibration Tracker" and "Calibration Summary"; the TESTs use CalTracker and CalSummary
# ASSUME: Calibration Summary already exists and is rebuilt on every run: a header row Category and Count, then three labelled rows
# ASSUME: overdue means due before today; due soon means due from today up to 30 days from today; fine means due later than that
# ASSUME: due dates are whole days; each row with a real date is one instrument
# ASSUME: a blank or non-date Next Calibration Due is not counted in any of the three
GOAL     Count the calibration instruments that are overdue, due within 30 days and fine, and write the three counts with labels on the Calibration Summary sheet.
PROC     SummariseCalibrationDue
KIND     Function
MODULE   modCalibration
ARGS     tracker As Worksheet, out As Worksheet
RET      Long
VAR      cDue As Long
VAR      overdue As Long
VAR      soon As Long
VAR      fine As Long
PRE      COL($tracker, "Next Calibration Due") > 0
MODIFIES $out
OP       LET(cDue, COL($tracker, "Next Calibration Due"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($tracker, cDue) var=r
OP         SKIP_IF NOT IS_DATE(CELL($tracker, r, cDue))
OP         IF CELL($tracker, r, cDue) < TODAY
OP           LET(overdue, overdue + 1)
OP         ELSE
OP           IF CELL($tracker, r, cDue) <= TODAY + 30
OP             LET(soon, soon + 1)
OP           ELSE
OP             LET(fine, fine + 1)
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Category", "Count")
OP       WRITE_ROW($out, 2, "Overdue", overdue)
OP       WRITE_ROW($out, 3, "Due in 30 days", soon)
OP       WRITE_ROW($out, 4, "Fine", fine)
OP       RETURN(overdue + soon + fine)
ERR      standard
FIXTURE  cal_tracker_01.xlsx sheet=CalTracker cols=A:Instrument ID,B:Type,C:Next Calibration Due rows=11 data=A:seq(101,1);B:values("Gauge"x6,"Caliper"x5);C:reldates(-60,90)+boundary(-1,0,30,31,"",-30,5,90,-60,"",15)
FIXTURE  cal_tracker_01.xlsx sheet=CalSummary cols=A:Category,B:Count rows=0 stale=2
TEST     fixture=cal_tracker_01.xlsx call=SummariseCalibrationDue(CalTracker, CalSummary) runs=2 assert=RESULT = 9 AND ERROR_COUNT = 0 AND ROWS_USED(CalSummary) = 4 AND COUNTIF(CalSummary![Category], "stale") = 0 AND VALUE(CalSummary![Category]2) = "Overdue" AND VALUE(CalSummary![Count]2) = 3 AND VALUE(CalSummary![Category]3) = "Due in 30 days" AND VALUE(CalSummary![Count]3) = 4 AND VALUE(CalSummary![Category]4) = "Fine" AND VALUE(CalSummary![Count]4) = 2
TEST     fixture=cal_tracker_01.xlsx call=SummariseCalibrationDue(CalTracker, CalSummary) assert=SUM(CalSummary![Count]) = COUNTA(CalTracker![Next Calibration Due]) - 1 AND COUNTA(CalTracker![Next Calibration Due]) = PRE_COUNTA(CalTracker![Next Calibration Due]) AND COUNTA(CalTracker![Instrument ID]) = PRE_COUNTA(CalTracker![Instrument ID])
TEST     fixture=cal_tracker_01.xlsx call=SummariseCalibrationDue(CalSummary, CalTracker) assert=RESULT = -1 AND ERROR_COUNT = 1
