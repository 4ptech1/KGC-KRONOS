VCL 0.6
BUILD
# ASSUME: NCR Register headers (row 1) include "NCR No", "Issue Date", "Status" and "ITP No"; Issue Date holds real dates
# ASSUME: ITP Tracker headers include "ITP No", "Point", "Inspection Status" and "Hold Flag"; the Hold Flag column already exists
# ASSUME: the instruments sit on a sheet Calibration Register with headers "Instrument ID" and "Next Calibration Due" (real dates)
# ASSUME: the real sheets are "NCR Register", "ITP Tracker", "Calibration Register", and the output sheets "Overdue NCRs", "Calibration Due" and "QA Summary" already exist and are rebuilt on every run
# ASSUME: the TESTs use the short sheet names Ncr, Itp, Cal, OverdueNcrs, CalDue and QaSummary
# ASSUME: "Open" is matched ignoring case and extra spaces; older than 30 days means issued before today minus 30 days
# ASSUME: overdue or due within 15 days means Next Calibration Due on or before today plus 15 days; an instrument with no valid date is not listed
# ASSUME: a point is held when any Open NCR (of any age) has its ITP No, matched after trimming and ignoring case; an Open NCR with a blank ITP No holds nothing
# ASSUME: Hold Flag is cleared only where it says NCR Hold; any other text there stays
# ASSUME: QA Summary has the columns Measure and Count, with the three measures in the order of the request; listed rows keep the order of their source sheet
GOAL     Rebuild the weekly QA report: list overdue Open NCRs, flag ITP points held by Open NCRs, list calibrations due, and summarise the three counts.
NAME     WeeklyQaReport
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modNcr
MODULE   modCalibration
MODULE   modReport
SIG      modMain.Run()
SIG      modMain.RunHeadless(ncr As Worksheet, itp As Worksheet, cal As Worksheet, overdue As Worksheet, caldue As Worksheet, summary As Worksheet) As Long
SIG      modNcr.ListOverdueNcrs(ncr As Worksheet, out As Worksheet) As Long
SIG      modNcr.FlagHeldPoints(ncr As Worksheet, itp As Worksheet) As Long
SIG      modCalibration.ListCalibrationDue(cal As Worksheet, out As Worksheet) As Long
SIG      modReport.WriteSummary(summary As Worksheet, nOverdue As Long, nHeld As Long, nDue As Long)
FLOW     modMain.Run -> modMain.RunHeadless -> modNcr.ListOverdueNcrs
FLOW     modMain.RunHeadless -> modNcr.FlagHeldPoints
FLOW     modMain.RunHeadless -> modCalibration.ListCalibrationDue
FLOW     modMain.RunHeadless -> modReport.WriteSummary
FIXTURE  qa_01.xlsx sheet=Ncr cols=A:NCR No,B:Issue Date,C:Status,D:ITP No rows=10 data=A:seq(1,1);B:reldates(-90,-1)+boundary(-31,-30,-29,-90,-60,"",-45,-40,-75,-5);C:values("Open"x7," Open"x1,"Closed"x2)+boundary("Open","Open","Open","Open","Closed","Open","Open"," Open","Closed","Open");D:values("ITP-1"x1,"ITP-2"x2,"ITP-3"x1,"ITP-4"x1,""x1,"  itp-3"x1,"ITP-5"x1,"ITP-6"x1,"ITP-7"x1)+boundary("ITP-1","ITP-2","ITP-2","ITP-3","ITP-4","","  itp-3","ITP-5","ITP-6","ITP-7")
FIXTURE  qa_01.xlsx sheet=Itp cols=A:ITP No,B:Point,C:Inspection Status,D:Hold Flag rows=12 data=A:values("ITP-1"x2,"ITP-2"x1,"ITP-3"x1," ITP-3 "x1,"ITP-4"x2,"ITP-5"x1,"ITP-6"x1,"ITP-7"x1,"ITP-8"x1,""x1)+boundary("ITP-4","ITP-1","ITP-2","","ITP-3","ITP-6","ITP-1","ITP-5","ITP-4","ITP-8"," ITP-3 ","ITP-7");B:seq(1,1);C:values("Pending"x6,"Passed"x6);D:values("NCR Hold"x6,"Witness"x1,""x5)+boundary("NCR Hold","","NCR Hold","NCR Hold","","Witness","NCR Hold","","NCR Hold","","","NCR Hold")
FIXTURE  qa_01.xlsx sheet=Cal cols=A:Instrument ID,B:Next Calibration Due rows=10 data=A:seq(101,1);B:reldates(-60,90)+boundary(-10,15,16,0,-1,"",14,90,-60,30)
FIXTURE  qa_01.xlsx sheet=OverdueNcrs cols=A:NCR No,B:Issue Date,C:Status,D:ITP No rows=0 stale=6
FIXTURE  qa_01.xlsx sheet=CalDue cols=A:Instrument ID,B:Next Calibration Due rows=0 stale=8
FIXTURE  qa_01.xlsx sheet=QaSummary cols=A:Measure,B:Count rows=0 stale=5
TEST     fixture=qa_01.xlsx call=modMain.RunHeadless(Ncr, Itp, Cal, OverdueNcrs, CalDue, QaSummary) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(OverdueNcrs![NCR No]) = 5 AND COUNTA(OverdueNcrs![Issue Date]) = 5 AND COUNTIF(OverdueNcrs![NCR No], "stale") = 0 AND FILL_COUNT(OverdueNcrs![NCR No], #FF0000) = 0 AND COUNTA(CalDue![Instrument ID]) = 7 AND COUNTA(CalDue![Next Calibration Due]) = 7 AND COUNTIF(CalDue![Instrument ID], "stale") = 0 AND FILL_COUNT(CalDue![Instrument ID], #FF0000) = 0 AND COUNTIF(Itp![Hold Flag], "NCR Hold") = 7 AND COUNTA(QaSummary![Measure]) = 4 AND VALUE(QaSummary![Count]2) = 4 AND VALUE(QaSummary![Count]3) = 7 AND VALUE(QaSummary![Count]4) = 6 AND COUNTIF(QaSummary![Measure], "stale") = 0 AND FILL_COUNT(QaSummary![Measure], #FF0000) = 0
TEST     fixture=qa_01.xlsx call=modMain.RunHeadless(Ncr, Itp, Cal, OverdueNcrs, CalDue, QaSummary) assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(Ncr![NCR No]) = PRE_COUNTA(Ncr![NCR No]) AND COUNTIF(Ncr![Status], "Open") = PRE_COUNTIF(Ncr![Status], "Open") AND COUNTA(Itp![ITP No]) = PRE_COUNTA(Itp![ITP No]) AND COUNTA(Cal![Instrument ID]) = PRE_COUNTA(Cal![Instrument ID]) AND VALUE(OverdueNcrs![NCR No]2) = 1 AND VALUE(OverdueNcrs![NCR No]3) = 4 AND VALUE(OverdueNcrs![NCR No]4) = 7 AND VALUE(OverdueNcrs![NCR No]5) = 8 AND VALUE(OverdueNcrs![ITP No]2) = "ITP-1" AND COUNTA(Itp![Hold Flag]) = 9 AND VALUE(Itp![Hold Flag]12) = "NCR Hold" AND VALUE(Itp![Hold Flag]7) = "Witness" AND VALUE(CalDue![Instrument ID]2) = 101 AND VALUE(CalDue![Instrument ID]7) = 109 AND VALUE(QaSummary![Measure]2) = "Overdue NCRs" AND VALUE(QaSummary![Measure]3) = "Held ITP points" AND VALUE(QaSummary![Measure]4) = "Instruments due"
TEST     fixture=qa_01.xlsx call=modMain.RunHeadless(Itp, Ncr, Cal, OverdueNcrs, CalDue, QaSummary) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
TASK
GOAL     Start the weekly QA report on the workbook's own sheets and tell the user how many Open NCRs are overdue.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "NCR Register"), SHEET(THIS_BOOK, "ITP Tracker"), SHEET(THIS_BOOK, "Calibration Register"), SHEET(THIS_BOOK, "Overdue NCRs"), SHEET(THIS_BOOK, "Calibration Due"), SHEET(THIS_BOOK, "QA Summary"), into=n)
OP       MSG("Overdue NCRs listed: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run the four steps of the weekly QA report on the sheets it is given and return how many overdue Open NCRs were listed.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     ncr As Worksheet, itp As Worksheet, cal As Worksheet, overdue As Worksheet, caldue As Worksheet, summary As Worksheet
RET      Long
VAR      nOverdue As Long
VAR      nHeld As Long
VAR      nDue As Long
PRE      COL($ncr, "NCR No") > 0; COL($ncr, "Issue Date") > 0; COL($ncr, "Status") > 0; COL($ncr, "ITP No") > 0; COL($itp, "ITP No") > 0; COL($itp, "Hold Flag") > 0; COL($cal, "Instrument ID") > 0; COL($cal, "Next Calibration Due") > 0
MODIFIES $itp["Hold Flag"]
MODIFIES $overdue
MODIFIES $caldue
MODIFIES $summary
OP       CALL(modNcr.ListOverdueNcrs, $ncr, $overdue, into=nOverdue)
OP       CALL(modNcr.FlagHeldPoints, $ncr, $itp, into=nHeld)
OP       CALL(modCalibration.ListCalibrationDue, $cal, $caldue, into=nDue)
OP       CALL(modReport.WriteSummary, $summary, nOverdue, nHeld, nDue)
OP       RETURN(nOverdue)
ERR      standard
---
VCL 0.6
TASK
# ASSUME: the output sheet is rebuilt from scratch on every run and gets the same four headers as the register
GOAL     Rebuild the Overdue NCRs sheet with every Open NCR issued more than 30 days ago and return how many were listed.
PROC     ListOverdueNcrs
KIND     Function
MODULE   modNcr
ARGS     ncr As Worksheet, out As Worksheet
RET      Long
VAR      cNo As Long
VAR      cDate As Long
VAR      cStatus As Long
VAR      cItp As Long
VAR      outRow As Long
PRE      COL($ncr, "NCR No") > 0; COL($ncr, "Issue Date") > 0; COL($ncr, "Status") > 0; COL($ncr, "ITP No") > 0
MODIFIES $out
OP       LET(cNo, COL($ncr, "NCR No"))
OP       LET(cDate, COL($ncr, "Issue Date"))
OP       LET(cStatus, COL($ncr, "Status"))
OP       LET(cItp, COL($ncr, "ITP No"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "NCR No", "Issue Date", "Status", "ITP No")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, cNo) var=r
OP         IF SAME(CELL($ncr, r, cStatus), "Open")
OP           IF IS_DATE(CELL($ncr, r, cDate))
OP             IF CELL($ncr, r, cDate) < TODAY - 30
OP               WRITE_ROW($out, outRow, CELL($ncr, r, cNo), CELL($ncr, r, cDate), CELL($ncr, r, cStatus), CELL($ncr, r, cItp))
OP               LET(outRow, outRow + 1)
OP       RETURN(outRow - 2)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write NCR Hold in the Hold Flag of every ITP point that has an Open NCR, clear that flag on the other points and return how many points are held.
PROC     FlagHeldPoints
KIND     Function
MODULE   modNcr
ARGS     ncr As Worksheet, itp As Worksheet
RET      Long
VAR      cStatus As Long
VAR      cNcrItp As Long
VAR      cItp As Long
VAR      cFlag As Long
VAR      openItps As Dictionary
VAR      held As Long
PRE      COL($ncr, "Status") > 0; COL($ncr, "ITP No") > 0; COL($itp, "ITP No") > 0; COL($itp, "Hold Flag") > 0
MODIFIES $itp["Hold Flag"]
OP       LET(cStatus, COL($ncr, "Status"))
OP       LET(cNcrItp, COL($ncr, "ITP No"))
OP       LET(cItp, COL($itp, "ITP No"))
OP       LET(cFlag, COL($itp, "Hold Flag"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, cStatus) var=r
OP         IF SAME(CELL($ncr, r, cStatus), "Open") AND NOT IS_EMPTY(CELL($ncr, r, cNcrItp))
OP           DICT_ADD(openItps, TRIM(CELL($ncr, r, cNcrItp)), 1, mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($itp, cItp) var=p
OP         IF EXISTS(openItps, TRIM(CELL($itp, p, cItp)))
OP           SET_VALUE(CELL($itp, p, cFlag), "NCR Hold")
OP           LET(held, held + 1)
OP         ELSE
OP           IF SAME(CELL($itp, p, cFlag), "NCR Hold")
OP             CLEAR(CELL($itp, p, cFlag), what=contents)
OP       RETURN(held)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Calibration Due sheet with every instrument that is overdue or due within 15 days and return how many were listed.
PROC     ListCalibrationDue
KIND     Function
MODULE   modCalibration
ARGS     cal As Worksheet, out As Worksheet
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
OP           IF CELL($cal, r, cDue) <= TODAY + 15
OP             WRITE_ROW($out, outRow, CELL($cal, r, cId), CELL($cal, r, cDue))
OP             LET(outRow, outRow + 1)
OP       RETURN(outRow - 2)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the QA Summary sheet with a label and a count for overdue NCRs, held ITP points and instruments due.
PROC     WriteSummary
KIND     Sub
MODULE   modReport
ARGS     summary As Worksheet, nOverdue As Long, nHeld As Long, nDue As Long
MODIFIES $summary
OP       CLEAR(USED_RANGE($summary), what=all)
OP       WRITE_ROW($summary, 1, "Measure", "Count")
OP       WRITE_ROW($summary, 2, "Overdue NCRs", $nOverdue)
OP       WRITE_ROW($summary, 3, "Held ITP points", $nHeld)
OP       WRITE_ROW($summary, 4, "Instruments due", $nDue)
ERR      standard
