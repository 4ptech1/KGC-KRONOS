VCL 0.1
TASK
# ASSUME: Calibration Tracker headers (row 1) include "Instrument ID" and "Next Calibration Due"
# ASSUME: an instrument is listed when its due date is blank or invalid, already passed, or within the next 15 days
# ASSUME: the Overdue sheet is rebuilt every run with the tracker's header row plus a Status column, soonest due first
# ASSUME: rows with a blank Instrument ID are skipped
PROC     BuildCalibrationDueList
KIND     Function
MODULE   modQuality
ARGS     cal As Worksheet, out As Worksheet
RET      Long
VAR      cId As Long
VAR      cDue As Long
VAR      lastCol As Long
VAR      outRow As Long
VAR      status As String
PRE      COL($cal, "Instrument ID") > 0; COL($cal, "Next Calibration Due") > 0
OP       LET(cId, COL($cal, "Instrument ID"))
OP       LET(cDue, COL($cal, "Next Calibration Due"))
OP       LET(lastCol, LASTCOL($cal, 1))
OP       CLEAR(USED_RANGE($out), what=all)
OP       COPY_ROW($cal, 1, $out, 1, cols=lastCol)
OP       SET_VALUE(CELL($out, 1, lastCol + 1), "Status")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($cal, cId) var=r
OP         IF NOT IS_EMPTY(CELL($cal, r, cId))
OP           LET(status, "")
OP           IF NOT IS_DATE(CELL($cal, r, cDue)) THEN LET(status, "Due date missing")
OP           IF IS_DATE(CELL($cal, r, cDue))
OP             IF CELL($cal, r, cDue) < TODAY THEN LET(status, "Overdue")
OP             IF CELL($cal, r, cDue) >= TODAY AND CELL($cal, r, cDue) <= TODAY + 15 THEN LET(status, "Due within 15 days")
OP           IF NOT IS_EMPTY(status)
OP             COPY_ROW($cal, r, $out, outRow, cols=lastCol)
OP             SET_VALUE(CELL($out, outRow, lastCol + 1), status)
OP             LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=cDue, order=asc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  calibration_01.xlsx sheet=Calibration cols=A:Instrument ID,B:Description,C:Last Calibrated,D:Next Calibration Due rows=30 data=A:seq(1,1);B:text(15);C:dates(2025-01-01,2026-09-30);D:reldates(-500,990)+boundary("",0,15,16,-1)+blank(2)
FIXTURE  calibration_01.xlsx sheet=Overdue cols=A:Instrument ID rows=0 stale=10
TEST     fixture=calibration_01.xlsx call=BuildCalibrationDueList(Calibration, Overdue) runs=2 assert=RESULT = 12 AND COUNTA(Overdue![Instrument ID]) = 13 AND COUNTIF(Overdue![Instrument ID], "stale") = 0 AND COUNTIF(Overdue![Status], "Overdue") = 7 AND COUNTIF(Overdue![Status], "Due within 15 days") = 2 AND ERROR_COUNT = 0
TEST     fixture=calibration_01.xlsx call=BuildCalibrationDueList(Calibration, Overdue) assert=COUNTIF(Overdue![Status], "Due date missing") = 3
TEST     fixture=calibration_01.xlsx call=BuildCalibrationDueList(Calibration, Overdue) assert=COUNTA(Calibration![Next Calibration Due]) = PRE_COUNTA(Calibration![Next Calibration Due]) AND COUNTA(Calibration![Instrument ID]) = PRE_COUNTA(Calibration![Instrument ID])
TEST     fixture=calibration_01.xlsx call=BuildCalibrationDueList(Overdue, Overdue) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Overdue![Instrument ID]) = PRE_COUNTA(Overdue![Instrument ID])
