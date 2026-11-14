VCL 0.6
BUILD
# ASSUME: the real sheets are "Calibration Tracker", "Due List" and "Location Summary", and all three exist; the TESTs use Tracker, DueList and Summary
# ASSUME: Calibration Tracker headers (row 1) include "Instrument ID", "Type", "Location" and "Next Calibration Due"
# ASSUME: overdue means a due date before today; due soon means from today to 30 days ahead, both ends included
# ASSUME: a due date that is blank or typed as text is not a date: that instrument is not coloured, not listed and counts as neither
# ASSUME: a row with a blank Instrument ID is not an instrument; rows are read down to the last Instrument ID
# ASSUME: old highlights are cleared from row 2 down, from column A to the last heading; the header row keeps its fill
# ASSUME: Due List and Location Summary are cleared and rebuilt on every run; nothing else may be kept on them
# ASSUME: Location Summary lists every location that has an instrument, also with 0 and 0; an instrument with no location counts under "(no location)"
# ASSUME: locations are matched ignoring case and spaces at the ends; the first spelling found is written
GOAL     Highlight overdue and due-soon instruments in the calibration tracker, rebuild a due list and a summary by location, and report how many are overdue.
NAME     CalibrationFollowUp
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(tracker As Worksheet, dueList As Worksheet, summary As Worksheet) As Long
SIG      modSteps.HighlightDueRows(tracker As Worksheet) As Long
SIG      modSteps.WriteDueList(tracker As Worksheet, out As Worksheet)
SIG      modSteps.WriteLocationSummary(tracker As Worksheet, out As Worksheet)
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.HighlightDueRows
FLOW     modMain.RunHeadless -> modSteps.WriteDueList
FLOW     modMain.RunHeadless -> modSteps.WriteLocationSummary
FIXTURE  cal_followup_01.xlsx sheet=Tracker cols=A:Instrument ID,B:Type,C:Location,D:Next Calibration Due rows=12 stale=3 data=A:seq(3001,1)+boundary(3001,3002,3003,3004,3005,"",3007,3008,3009,3010,3011,3012);B:values("Gauge"x6,"Caliper"x5,"Torque wrench"x1)+boundary("Gauge","Gauge","Gauge","Gauge","Gauge","Gauge","Caliper","Caliper","Torque wrench");C:values("Bay 1"x3,"Bay 2"x2,"Stores"x3,"bay 1 "x1,""x1,"Annex"x1," BAY 2"x1)+boundary("Bay 1","Bay 2","Bay 1","Stores","bay 1 ","Bay 2","","Annex","Stores"," BAY 2","Stores","Bay 1");D:reldates(-60,90)+boundary(-1,0,30,31,"",-5,-10,90,-40,12,-20,5)
FIXTURE  cal_followup_01.xlsx sheet=DueList cols=A:Instrument ID,B:Type,C:Location,D:Next Calibration Due rows=0 stale=10
FIXTURE  cal_followup_01.xlsx sheet=Summary cols=A:Location,B:Overdue,C:Due Soon rows=0 stale=7
FIXTURE  cal_followup_02.xlsx sheet=Tracker cols=A:Instrument ID,B:Type,C:Location,D:Next Calibration Due rows=5 stale=2 data=A:seq(4001,1);B:values("Gauge"x5);C:values("Bay 1"x3,"Stores"x2);D:reldates(40,200)+boundary(31,"")
FIXTURE  cal_followup_02.xlsx sheet=DueList cols=A:Instrument ID,B:Type,C:Location,D:Next Calibration Due rows=0 stale=2
FIXTURE  cal_followup_02.xlsx sheet=Summary cols=A:Location,B:Overdue,C:Due Soon rows=0 stale=4
TEST     fixture=cal_followup_01.xlsx call=modMain.RunHeadless(Tracker, DueList, Summary) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND FILL_COUNT(Tracker![Instrument ID], #FF0000) = 4 AND FILL_COUNT(Tracker![Instrument ID], #FFFF00) = 4 AND FILL_COUNT(Tracker![Next Calibration Due], #FF0000) = 4 AND FILL_COUNT(Tracker![Next Calibration Due], #FFFF00) = 4 AND FILL_COUNT(Tracker![Instrument ID]2, #FF0000) = 1 AND FILL_COUNT(Tracker![Instrument ID]3, #FFFF00) = 1 AND FILL_COUNT(Tracker![Instrument ID]4, #FFFF00) = 1 AND FILL_COUNT(Tracker![Instrument ID]13, #FFFF00) = 1 AND ROWS_USED(DueList) = 9 AND COUNTIF(DueList![Instrument ID], "stale") = 0 AND FILL_COUNT(DueList![Instrument ID], #FF0000) = 0 AND ROWS_USED(Summary) = 6 AND COUNTIF(Summary![Location], "stale") = 0 AND FILL_COUNT(Summary![Location], #FF0000) = 0
TEST     fixture=cal_followup_01.xlsx call=modMain.RunHeadless(Tracker, DueList, Summary) assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(Tracker![Instrument ID]) = PRE_COUNTA(Tracker![Instrument ID]) AND COUNTA(Tracker![Next Calibration Due]) = PRE_COUNTA(Tracker![Next Calibration Due]) AND VALUE(DueList![Instrument ID]2) = 3009 AND VALUE(DueList![Instrument ID]3) = 3011 AND VALUE(DueList![Instrument ID]4) = 3007 AND VALUE(DueList![Instrument ID]5) = 3001 AND VALUE(DueList![Instrument ID]6) = 3002 AND VALUE(DueList![Instrument ID]7) = 3012 AND VALUE(DueList![Instrument ID]8) = 3010 AND VALUE(DueList![Instrument ID]9) = 3003 AND VALUE(DueList![Type]2) = "Torque wrench" AND COUNTIF(DueList![Type], "Torque wrench") = 1 AND VALUE(DueList![Location]2) = "Stores" AND VALUE(DueList![Location]5) = "Bay 1" AND VALUE(DueList![Next Calibration Due]2) <= VALUE(DueList![Next Calibration Due]3) AND VALUE(DueList![Next Calibration Due]1) = "Next Calibration Due"
TEST     fixture=cal_followup_01.xlsx call=modMain.RunHeadless(Tracker, DueList, Summary) assert=ERROR_COUNT = 0 AND COUNTA(Tracker![Location]) = PRE_COUNTA(Tracker![Location]) AND VALUE(Summary![Location]2) = "(no location)" AND VALUE(Summary![Overdue]2) = 1 AND VALUE(Summary![Due Soon]2) = 0 AND VALUE(Summary![Location]3) = "Annex" AND VALUE(Summary![Overdue]3) = 0 AND VALUE(Summary![Due Soon]3) = 0 AND VALUE(Summary![Location]4) = "Bay 1" AND VALUE(Summary![Overdue]4) = 1 AND VALUE(Summary![Due Soon]4) = 2 AND VALUE(Summary![Location]5) = "Bay 2" AND VALUE(Summary![Overdue]5) = 0 AND VALUE(Summary![Due Soon]5) = 2 AND VALUE(Summary![Location]6) = "Stores" AND VALUE(Summary![Overdue]6) = 2 AND VALUE(Summary![Due Soon]6) = 0
TEST     fixture=cal_followup_02.xlsx call=modMain.RunHeadless(Tracker, DueList, Summary) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND ROWS_USED(DueList) = 1 AND VALUE(DueList![Next Calibration Due]1) = "Next Calibration Due" AND ROWS_USED(Summary) = 3 AND VALUE(Summary![Location]2) = "Bay 1" AND VALUE(Summary![Overdue]2) = 0 AND VALUE(Summary![Due Soon]2) = 0 AND VALUE(Summary![Location]3) = "Stores" AND FILL_COUNT(Tracker![Instrument ID], #FF0000) = 0 AND FILL_COUNT(Tracker![Instrument ID], #FFFF00) = 0 AND COUNTA(Tracker![Instrument ID]) = PRE_COUNTA(Tracker![Instrument ID])
TEST     fixture=cal_followup_01.xlsx call=modMain.RunHeadless(Summary, DueList, Tracker) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
TASK
GOAL     Start the calibration follow-up on the workbook's own sheets and tell the user how many instruments are overdue.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "Calibration Tracker"), SHEET(THIS_BOOK, "Due List"), SHEET(THIS_BOOK, "Location Summary"), into=n)
OP       IF n < 0
OP         MSG("The calibration follow-up did not run: check the headings of the Calibration Tracker.")
OP       ELSE
OP         MSG("Instruments overdue for calibration: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run the three steps of the calibration follow-up on the sheets it is given and return how many instruments are overdue.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     tracker As Worksheet, dueList As Worksheet, summary As Worksheet
RET      Long
VAR      overdue As Long
PRE      COL($tracker, "Instrument ID") > 0; COL($tracker, "Type") > 0; COL($tracker, "Location") > 0; COL($tracker, "Next Calibration Due") > 0
MODIFIES $tracker fills
MODIFIES $dueList
MODIFIES $summary
OP       CALL(modSteps.HighlightDueRows, $tracker, into=overdue)
OP       CALL(modSteps.WriteDueList, $tracker, $dueList)
OP       CALL(modSteps.WriteLocationSummary, $tracker, $summary)
OP       RETURN(overdue)
ERR      standard
---
VCL 0.6
TASK
GOAL     Clear old row highlights in the tracker, colour overdue instrument rows red and rows due within 30 days yellow, and return how many are overdue.
PROC     HighlightDueRows
KIND     Function
MODULE   modSteps
ARGS     tracker As Worksheet
RET      Long
VAR      cId As Long
VAR      cDue As Long
VAR      lastHead As Long
VAR      overdue As Long
PRE      COL($tracker, "Instrument ID") > 0; COL($tracker, "Next Calibration Due") > 0
MODIFIES $tracker fills
OP       LET(cId, COL($tracker, "Instrument ID"))
OP       LET(cDue, COL($tracker, "Next Calibration Due"))
OP       LET(lastHead, LASTCOL($tracker, 1))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($tracker, cId) var=r
OP         FORMAT(ROW_RANGE($tracker, r, 1, lastHead), fill=none)
OP         SKIP_IF IS_EMPTY(CELL($tracker, r, cId))
OP         IF IS_DATE(CELL($tracker, r, cDue))
OP           IF CELL($tracker, r, cDue) < TODAY
OP             FORMAT(ROW_RANGE($tracker, r, 1, lastHead), fill=red)
OP             LET(overdue, overdue + 1)
OP           ELSE
OP             IF CELL($tracker, r, cDue) <= TODAY + 30 THEN FORMAT(ROW_RANGE($tracker, r, 1, lastHead), fill=yellow)
OP       RETURN(overdue)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the due list with every instrument that is overdue or due within 30 days, sorted with the earliest due date first.
PROC     WriteDueList
KIND     Sub
MODULE   modSteps
ARGS     tracker As Worksheet, out As Worksheet
VAR      cId As Long
VAR      cType As Long
VAR      cLoc As Long
VAR      cDue As Long
VAR      outRow As Long
PRE      COL($tracker, "Instrument ID") > 0; COL($tracker, "Type") > 0; COL($tracker, "Location") > 0; COL($tracker, "Next Calibration Due") > 0
MODIFIES $out
OP       LET(cId, COL($tracker, "Instrument ID"))
OP       LET(cType, COL($tracker, "Type"))
OP       LET(cLoc, COL($tracker, "Location"))
OP       LET(cDue, COL($tracker, "Next Calibration Due"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Instrument ID", "Type", "Location", "Next Calibration Due")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($tracker, cId) var=r
OP         SKIP_IF IS_EMPTY(CELL($tracker, r, cId))
OP         IF IS_DATE(CELL($tracker, r, cDue))
OP           IF CELL($tracker, r, cDue) <= TODAY + 30
OP             WRITE_ROW($out, outRow, CELL($tracker, r, cId), CELL($tracker, r, cType), CELL($tracker, r, cLoc), CELL($tracker, r, cDue))
OP             LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=4, order=asc, header=yes)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the location summary with one row for each location and how many of its instruments are overdue and due soon, sorted A to Z.
PROC     WriteLocationSummary
KIND     Sub
MODULE   modSteps
ARGS     tracker As Worksheet, out As Worksheet
VAR      cId As Long
VAR      cLoc As Long
VAR      cDue As Long
VAR      place As String
VAR      isOver As Long
VAR      isSoon As Long
VAR      overByPlace As Dictionary
VAR      soonByPlace As Dictionary
VAR      nOver As Long
VAR      nSoon As Long
VAR      outRow As Long
PRE      COL($tracker, "Instrument ID") > 0; COL($tracker, "Location") > 0; COL($tracker, "Next Calibration Due") > 0
MODIFIES $out
OP       LET(cId, COL($tracker, "Instrument ID"))
OP       LET(cLoc, COL($tracker, "Location"))
OP       LET(cDue, COL($tracker, "Next Calibration Due"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Location", "Overdue", "Due Soon")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($tracker, cId) var=r
OP         SKIP_IF IS_EMPTY(CELL($tracker, r, cId))
OP         LET(place, TRIM(CELL($tracker, r, cLoc)))
OP         IF place = "" THEN LET(place, "(no location)")
OP         LET(isOver, 0)
OP         LET(isSoon, 0)
OP         IF IS_DATE(CELL($tracker, r, cDue))
OP           IF CELL($tracker, r, cDue) < TODAY
OP             LET(isOver, 1)
OP           ELSE
OP             IF CELL($tracker, r, cDue) <= TODAY + 30 THEN LET(isSoon, 1)
OP         DICT_ADD(overByPlace, place, isOver, mode=sum)
OP         DICT_ADD(soonByPlace, place, isSoon, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(overByPlace)
OP         DICT_GET(overByPlace, k, into=nOver)
OP         DICT_GET(soonByPlace, k, into=nSoon)
OP         WRITE_ROW($out, outRow, k, nOver, nSoon)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
ERR      standard
