VCL 0.6
BUILD
# ASSUME: the real sheets are "Consumable Stock", "Reorder List", "Calibration Tracker" and "Weekly Summary"; the TESTs use Stock, Reorder, Calib and Summary
# ASSUME: Reorder List and Weekly Summary already exist; both are cleared and rebuilt on every run
# ASSUME: Consumable Stock headers (row 1) include "Batch No", "Qty On Hand", "Reorder Level" and "Remark"; Calibration Tracker headers include "Instrument ID" and "Next Calibration Due"
# ASSUME: a stock line needs reordering when Qty On Hand and Reorder Level are both numbers and Qty On Hand is below Reorder Level; the Remark of every other line is cleared
# ASSUME: stock lines run from row 2 to the last Batch No, instruments from row 2 to the last Instrument ID
# ASSUME: overdue means Next Calibration Due is before today; due soon means from today to 15 days ahead; a blank or text due date gets no fill and is not counted
# ASSUME: only the Instrument ID cells (row 2 down) lose their old fills; other fills on the Calibration Tracker are kept
# ASSUME: the Reorder List gets the values (not the formats) of every column that has a heading in row 1 of Consumable Stock
GOAL     Mark stock lines to reorder, rebuild the reorder list, highlight overdue and soon-due calibrations and rebuild the weekly summary.
NAME     WeeklyShopTool
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(stock As Worksheet, reorder As Worksheet, cal As Worksheet, summary As Worksheet) As Long
SIG      modSteps.MarkReorder(stock As Worksheet) As Long
SIG      modSteps.CopyReorderRows(stock As Worksheet, reorder As Worksheet)
SIG      modSteps.FlagCalibration(cal As Worksheet, soonDays As Long)
SIG      modSteps.WriteSummary(stock As Worksheet, cal As Worksheet, summary As Worksheet, soonDays As Long)
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.MarkReorder
FLOW     modMain.RunHeadless -> modSteps.CopyReorderRows
FLOW     modMain.RunHeadless -> modSteps.FlagCalibration
FLOW     modMain.RunHeadless -> modSteps.WriteSummary
FIXTURE  shop_01.xlsx sheet=Stock cols=A:Batch No,B:Qty On Hand,C:Reorder Level,D:Remark,E:Location rows=8 stale=2 data=A:seq(101,1);B:int(0,40)+boundary(9,10,11,"",0,-2,30,4);C:int(1,30)+boundary(10,10,10,10,1,"",25,12);D:values("Reorder"x4,"old"x1,""x3)+boundary("","Reorder","old","Reorder","Reorder","Reorder","","");E:values("Bay 1"x3,"Bay 2"x3,""x2)+boundary("")
FIXTURE  shop_01.xlsx sheet=Reorder cols=A:Old List rows=0 stale=6
FIXTURE  shop_01.xlsx sheet=Calib cols=A:Instrument ID,B:Next Calibration Due rows=9 stale=9 data=A:seq(201,1);B:reldates(-30,40)+boundary(0,-1,15,16,"",-20,40,3,7)
FIXTURE  shop_01.xlsx sheet=Summary cols=A:Item,B:Count rows=0 stale=6
TEST     fixture=shop_01.xlsx call=modMain.RunHeadless(Stock, Reorder, Calib, Summary) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Stock![Remark], "Reorder") = 3 AND COUNTA(Stock![Remark]) = 4 AND COUNTA(Reorder!A:A) = 4 AND COUNTIF(Reorder!D:D, "Reorder") = 3 AND COUNTIF(Reorder!A:A, "stale") = 0 AND FILL_COUNT(Reorder!A1:A10, #FF0000) = 0 AND FILL_COUNT(Calib![Instrument ID], #FF0000) = 2 AND FILL_COUNT(Calib![Instrument ID], #FFFF00) = 4 AND VALUE(Summary!B2) = 3 AND VALUE(Summary!B3) = 2 AND VALUE(Summary!B4) = 4 AND COUNTA(Summary!A:A) = 4 AND COUNTIF(Summary!A:A, "stale") = 0 AND FILL_COUNT(Summary!A1:A10, #FF0000) = 0
TEST     fixture=shop_01.xlsx call=modMain.RunHeadless(Stock, Reorder, Calib, Summary) assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTA(Stock![Batch No]) = PRE_COUNTA(Stock![Batch No]) AND COUNTA(Calib![Next Calibration Due]) = PRE_COUNTA(Calib![Next Calibration Due]) AND VALUE(Stock![Remark]2) = "Reorder" AND VALUE(Stock![Remark]9) = "Reorder" AND VALUE(Reorder!A1) = "Batch No" AND VALUE(Reorder!E1) = "Location" AND VALUE(Reorder!A2) = 101 AND VALUE(Reorder!A3) = 105 AND VALUE(Reorder!C4) = 12 AND VALUE(Summary!A1) = "Item" AND VALUE(Summary!B1) = "Count" AND VALUE(Summary!A2) = "Reorder lines" AND VALUE(Summary!A3) = "Overdue instruments" AND VALUE(Summary!A4) = "Instruments due soon" AND FILL_COUNT(Calib![Instrument ID]2, #FFFF00) = 1 AND FILL_COUNT(Calib![Instrument ID]3, #FF0000) = 1
TEST     fixture=shop_01.xlsx call=modMain.RunHeadless(Calib, Reorder, Stock, Summary) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
TASK
GOAL     Start the weekly shop tool on the workbook's own four sheets and tell the user how many stock lines need reordering.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "Consumable Stock"), SHEET(THIS_BOOK, "Reorder List"), SHEET(THIS_BOOK, "Calibration Tracker"), SHEET(THIS_BOOK, "Weekly Summary"), into=n)
OP       IF n < 0
OP         MSG("The weekly shop tool stopped: see ForgeLog.")
OP       ELSE
OP         MSG("Stock lines to reorder: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run the four weekly steps on the stock, reorder list, calibration and summary sheets it is given and return how many stock lines need reordering.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     stock As Worksheet, reorder As Worksheet, cal As Worksheet, summary As Worksheet
RET      Long
VAR      marked As Long
PRE      COL($stock, "Batch No") > 0; COL($stock, "Qty On Hand") > 0; COL($stock, "Reorder Level") > 0; COL($stock, "Remark") > 0; COL($cal, "Instrument ID") > 0; COL($cal, "Next Calibration Due") > 0
MODIFIES $stock["Remark"]
MODIFIES $reorder
MODIFIES $cal fills
MODIFIES $summary
OP       CALL(modSteps.MarkReorder, $stock, into=marked)
OP       CALL(modSteps.CopyReorderRows, $stock, $reorder)
OP       CALL(modSteps.FlagCalibration, $cal, 15)
OP       CALL(modSteps.WriteSummary, $stock, $cal, $summary, 15)
OP       RETURN(marked)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Reorder in the Remark of every stock line whose quantity on hand is below its reorder level, clear the other remarks and return the count.
PROC     MarkReorder
KIND     Function
MODULE   modSteps
ARGS     stock As Worksheet
RET      Long
VAR      cBatch As Long
VAR      cQty As Long
VAR      cLevel As Long
VAR      cRemark As Long
VAR      n As Long
PRE      COL($stock, "Batch No") > 0; COL($stock, "Qty On Hand") > 0; COL($stock, "Reorder Level") > 0; COL($stock, "Remark") > 0
MODIFIES $stock["Remark"]
OP       LET(cBatch, COL($stock, "Batch No"))
OP       LET(cQty, COL($stock, "Qty On Hand"))
OP       LET(cLevel, COL($stock, "Reorder Level"))
OP       LET(cRemark, COL($stock, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($stock, cBatch) var=r
OP         CLEAR(CELL($stock, r, cRemark), what=contents)
OP         IF IS_NUMBER(CELL($stock, r, cQty))
OP           IF IS_NUMBER(CELL($stock, r, cLevel))
OP             IF CELL($stock, r, cQty) < CELL($stock, r, cLevel)
OP               SET_VALUE(CELL($stock, r, cRemark), "Reorder")
OP               LET(n, n + 1)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the reorder list sheet with the stock headings and a copy of every stock line whose Remark says Reorder.
PROC     CopyReorderRows
KIND     Sub
MODULE   modSteps
ARGS     stock As Worksheet, reorder As Worksheet
VAR      cBatch As Long
VAR      cRemark As Long
VAR      lastCol As Long
VAR      outRow As Long
PRE      COL($stock, "Batch No") > 0; COL($stock, "Remark") > 0
MODIFIES $reorder
OP       LET(cBatch, COL($stock, "Batch No"))
OP       LET(cRemark, COL($stock, "Remark"))
OP       LET(lastCol, LASTCOL($stock, 1))
OP       CLEAR(USED_RANGE($reorder), what=all)
OP       COPY_ROW($stock, 1, $reorder, 1, cols=lastCol)
OP       LET(outRow, 1)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($stock, cBatch) var=r
OP         IF SAME(CELL($stock, r, cRemark), "Reorder")
OP           LET(outRow, outRow + 1)
OP           COPY_ROW($stock, r, $reorder, outRow, cols=lastCol)
ERR      standard
---
VCL 0.6
TASK
GOAL     Clear the old fills on the Instrument ID cells, then fill each one red when its calibration is overdue or yellow when it is due within the given days.
PROC     FlagCalibration
KIND     Sub
MODULE   modSteps
ARGS     cal As Worksheet, soonDays As Long
VAR      cId As Long
VAR      cDue As Long
PRE      COL($cal, "Instrument ID") > 0; COL($cal, "Next Calibration Due") > 0
MODIFIES $cal fills
OP       LET(cId, COL($cal, "Instrument ID"))
OP       LET(cDue, COL($cal, "Next Calibration Due"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($cal, cId) var=r
OP         FORMAT(CELL($cal, r, cId), fill=none)
OP         IF IS_DATE(CELL($cal, r, cDue))
OP           IF CELL($cal, r, cDue) < TODAY
OP             FORMAT(CELL($cal, r, cId), fill=red)
OP           ELSE
OP             IF CELL($cal, r, cDue) <= TODAY + $soonDays THEN FORMAT(CELL($cal, r, cId), fill=yellow)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the weekly summary sheet with labelled counts of stock lines to reorder, overdue instruments and instruments due soon.
PROC     WriteSummary
KIND     Sub
MODULE   modSteps
ARGS     stock As Worksheet, cal As Worksheet, summary As Worksheet, soonDays As Long
VAR      cBatch As Long
VAR      cRemark As Long
VAR      cId As Long
VAR      cDue As Long
VAR      nReorder As Long
VAR      nOver As Long
VAR      nSoon As Long
PRE      COL($stock, "Batch No") > 0; COL($stock, "Remark") > 0; COL($cal, "Instrument ID") > 0; COL($cal, "Next Calibration Due") > 0
MODIFIES $summary
OP       LET(cBatch, COL($stock, "Batch No"))
OP       LET(cRemark, COL($stock, "Remark"))
OP       LET(cId, COL($cal, "Instrument ID"))
OP       LET(cDue, COL($cal, "Next Calibration Due"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($stock, cBatch) var=r
OP         IF SAME(CELL($stock, r, cRemark), "Reorder") THEN LET(nReorder, nReorder + 1)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($cal, cId) var=i
OP         IF IS_DATE(CELL($cal, i, cDue))
OP           IF CELL($cal, i, cDue) < TODAY
OP             LET(nOver, nOver + 1)
OP           ELSE
OP             IF CELL($cal, i, cDue) <= TODAY + $soonDays THEN LET(nSoon, nSoon + 1)
OP       CLEAR(USED_RANGE($summary), what=all)
OP       WRITE_ROW($summary, 1, "Item", "Count")
OP       WRITE_ROW($summary, 2, "Reorder lines", nReorder)
OP       WRITE_ROW($summary, 3, "Overdue instruments", nOver)
OP       WRITE_ROW($summary, 4, "Instruments due soon", nSoon)
ERR      standard
