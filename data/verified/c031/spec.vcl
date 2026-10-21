VCL 0.6
BUILD
# ASSUME: NCR Register headers (row 1) include "NCR No", "Issue Date" and "Status"; Calibration Tracker headers include "Instrument ID" and "Next Calibration Due"
# ASSUME: the real sheets are "NCR Register", "Calibration Tracker", "Overdue NCRs", "Calibration Due" and "Weekly Summary" and all five exist; the TESTs call them NCR, Calib, OverdueNcr, CalDue and Summary
# ASSUME: an NCR is listed when its Status is Open (ignoring case and spaces) and its Issue Date is more than 30 days before today; a row with no real date is not listed
# ASSUME: an instrument is listed when its Next Calibration Due is today, in the past, or at most 15 days ahead; a row with no real date is not listed
# ASSUME: the three output sheets are cleared and rebuilt on every run; the summary has the labels "Overdue NCRs" and "Calibration due" in column A and the counts in column B
GOAL     Rebuild the Overdue NCRs and Calibration Due lists and a Weekly Summary sheet with the count of each, all from the NCR Register and the Calibration Tracker.
NAME     WeeklyQuality
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(ncr As Worksheet, cal As Worksheet, ncrOut As Worksheet, calOut As Worksheet, summary As Worksheet) As Long
SIG      modSteps.ListOverdueNcrs(ncr As Worksheet, out As Worksheet, limit As Long) As Long
SIG      modSteps.ListCalibrationDue(cal As Worksheet, out As Worksheet, window As Long) As Long
SIG      modSteps.WriteWeeklySummary(summary As Worksheet, ncrCount As Long, calCount As Long)
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.ListOverdueNcrs
FLOW     modMain.RunHeadless -> modSteps.ListCalibrationDue
FLOW     modMain.RunHeadless -> modSteps.WriteWeeklySummary
FIXTURE  weekly_01.xlsx sheet=NCR cols=A:NCR No,B:Issue Date,C:Status rows=8 data=A:seq(101,1);B:reldates(-120,-1)+boundary(-31,-30,-29,-45,-60,"",-100,-31);C:values("Open"x4," open "x1,"OPEN"x1,"Closed"x2)+boundary("Open","Open","Open"," open ","Closed","Open","OPEN","Closed")
FIXTURE  weekly_01.xlsx sheet=Calib cols=A:Instrument ID,B:Next Calibration Due rows=8 data=A:seq(501,1);B:reldates(-30,90)+boundary(-5,0,14,15,16,90,"",-1)
FIXTURE  weekly_01.xlsx sheet=OverdueNcr cols=A:NCR No,B:Issue Date,C:Status rows=0 stale=2
FIXTURE  weekly_01.xlsx sheet=CalDue cols=A:Instrument ID,B:Next Calibration Due rows=0 stale=2
FIXTURE  weekly_01.xlsx sheet=Summary cols=A:Measure,B:Count rows=0 stale=2
TEST     fixture=weekly_01.xlsx call=modMain.RunHeadless(NCR, Calib, OverdueNcr, CalDue, Summary) runs=2 assert=RESULT = 8 AND ERROR_COUNT = 0 AND COUNTA(OverdueNcr![NCR No]) = 4 AND COUNTA(CalDue![Instrument ID]) = 6 AND COUNTA(Summary![Measure]) = 3 AND VALUE(Summary![Count]2) = 3 AND VALUE(Summary![Count]3) = 5 AND COUNTIF(OverdueNcr![NCR No], "stale") = 0 AND COUNTIF(CalDue![Instrument ID], "stale") = 0 AND COUNTIF(Summary![Measure], "stale") = 0
TEST     fixture=weekly_01.xlsx call=modMain.RunHeadless(NCR, Calib, OverdueNcr, CalDue, Summary) assert=RESULT = 8 AND COUNTA(NCR![NCR No]) = PRE_COUNTA(NCR![NCR No]) AND COUNTIF(NCR![Status], "Open") = PRE_COUNTIF(NCR![Status], "Open") AND COUNTA(Calib![Instrument ID]) = PRE_COUNTA(Calib![Instrument ID]) AND VALUE(OverdueNcr![NCR No]2) = 101 AND VALUE(OverdueNcr![NCR No]3) = 104 AND VALUE(OverdueNcr![NCR No]4) = 107 AND VALUE(OverdueNcr![Status]3) = " open " AND VALUE(CalDue![Instrument ID]2) = 501 AND VALUE(CalDue![Instrument ID]5) = 504 AND VALUE(CalDue![Instrument ID]6) = 508 AND VALUE(Summary![Measure]2) = "Overdue NCRs" AND VALUE(Summary![Measure]3) = "Calibration due"
TEST     fixture=weekly_01.xlsx call=modMain.RunHeadless(Summary, Calib, OverdueNcr, CalDue, NCR) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Summary![Measure]) = PRE_COUNTA(Summary![Measure])
---
VCL 0.6
TASK
GOAL     Start the weekly quality tool on the workbook's own sheets and tell the user how many NCRs and instruments were listed.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "NCR Register"), SHEET(THIS_BOOK, "Calibration Tracker"), SHEET(THIS_BOOK, "Overdue NCRs"), SHEET(THIS_BOOK, "Calibration Due"), SHEET(THIS_BOOK, "Weekly Summary"), into=n)
OP       MSG("Items listed this week: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run the whole weekly quality tool on the five sheets it is given and return how many NCRs and instruments were listed in total.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     ncr As Worksheet, cal As Worksheet, ncrOut As Worksheet, calOut As Worksheet, summary As Worksheet
RET      Long
VAR      nNcr As Long
VAR      nCal As Long
PRE      COL($ncr, "NCR No") > 0; COL($ncr, "Issue Date") > 0; COL($ncr, "Status") > 0; COL($cal, "Instrument ID") > 0; COL($cal, "Next Calibration Due") > 0
MODIFIES $ncrOut
MODIFIES $calOut
MODIFIES $summary
OP       CALL(modSteps.ListOverdueNcrs, $ncr, $ncrOut, 30, into=nNcr)
OP       CALL(modSteps.ListCalibrationDue, $cal, $calOut, 15, into=nCal)
OP       CALL(modSteps.WriteWeeklySummary, $summary, nNcr, nCal)
OP       RETURN(nNcr + nCal)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Overdue NCRs sheet with every Open NCR whose issue date is more than the limit of days ago, and return how many were listed.
PROC     ListOverdueNcrs
KIND     Function
MODULE   modSteps
ARGS     ncr As Worksheet, out As Worksheet, limit As Long
RET      Long
VAR      cNo As Long
VAR      cDate As Long
VAR      cStatus As Long
VAR      outRow As Long
PRE      COL($ncr, "NCR No") > 0; COL($ncr, "Issue Date") > 0; COL($ncr, "Status") > 0
MODIFIES $out
OP       LET(cNo, COL($ncr, "NCR No"))
OP       LET(cDate, COL($ncr, "Issue Date"))
OP       LET(cStatus, COL($ncr, "Status"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "NCR No", "Issue Date", "Status")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, cNo) var=r
OP         IF SAME(CELL($ncr, r, cStatus), "Open")
OP           IF IS_DATE(CELL($ncr, r, cDate))
OP             IF CELL($ncr, r, cDate) < TODAY - $limit
OP               WRITE_ROW($out, outRow, CELL($ncr, r, cNo), CELL($ncr, r, cDate), CELL($ncr, r, cStatus))
OP               LET(outRow, outRow + 1)
OP       RETURN(outRow - 2)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Calibration Due sheet with every instrument due within the window of days or already overdue, and return how many were listed.
PROC     ListCalibrationDue
KIND     Function
MODULE   modSteps
ARGS     cal As Worksheet, out As Worksheet, window As Long
RET      Long
VAR      cId As Long
VAR      cDue As Long
VAR      outRow As Long
PRE      COL($cal, "Instrument ID") > 0; COL($cal, "Next Calibration Due") > 0
MODIFIES $out
OP       LET(cId, COL($cal, "Instrument ID"))
OP       LET(cDue, COL($cal, "Next Calibration Due"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Instrument ID", "Next Calibration Due")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($cal, cId) var=r
OP         IF IS_DATE(CELL($cal, r, cDue))
OP           IF CELL($cal, r, cDue) <= TODAY + $window
OP             WRITE_ROW($out, outRow, CELL($cal, r, cId), CELL($cal, r, cDue))
OP             LET(outRow, outRow + 1)
OP       RETURN(outRow - 2)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Weekly Summary sheet with a label and a count for the NCRs and for the instruments that were listed.
PROC     WriteWeeklySummary
KIND     Sub
MODULE   modSteps
ARGS     summary As Worksheet, ncrCount As Long, calCount As Long
MODIFIES $summary
OP       CLEAR(USED_RANGE($summary), what=all)
OP       WRITE_ROW($summary, 1, "Measure", "Count")
OP       WRITE_ROW($summary, 2, "Overdue NCRs", $ncrCount)
OP       WRITE_ROW($summary, 3, "Calibration due", $calCount)
ERR      standard
