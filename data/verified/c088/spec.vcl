VCL 0.6
BUILD
# ASSUME: the real sheets are "Consumable Stock", "Reorder List" and "Class Summary" and all three exist; the TESTs use Stock, Reorder and Summary
# ASSUME: Consumable Stock headers in row 1 include "Batch No", "Filler Class", "Qty On Hand", "Reorder Level", "Expiry Date" and "Remark"; every batch has a Batch No
# ASSUME: a batch that expires today is not yet expired; a blank Expiry Date or one typed as text is never Expired
# ASSUME: Reorder needs Qty On Hand and Reorder Level to both be numbers; when either is not, the Remark is cleared unless the batch is expired
# ASSUME: Reorder List and Class Summary hold nothing else: both are cleared and rebuilt on every run; Class Summary gets the headings Filler Class, Total Qty On Hand, Expired Batches
# ASSUME: Filler Classes are matched ignoring case and spaces at the ends; a batch with no Filler Class is left out of Class Summary; a Qty On Hand that is not a number adds nothing to a total
# ASSUME: Reorder batches with the same Qty On Hand may come in either order; Run shows the number of expired batches in a message
GOAL     Mark each consumable batch as Expired or Reorder, rebuild the Reorder List and the Class Summary sheets and report how many batches are expired.
NAME     ConsumableStockReview
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(stock As Worksheet, reorder As Worksheet, summary As Worksheet) As Long
SIG      modSteps.MarkStockRemarks(stock As Worksheet) As Long
SIG      modSteps.ListReorderBatches(stock As Worksheet, reorder As Worksheet)
SIG      modSteps.SummariseFillerClasses(stock As Worksheet, summary As Worksheet)
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.MarkStockRemarks
FLOW     modMain.RunHeadless -> modSteps.ListReorderBatches
FLOW     modMain.RunHeadless -> modSteps.SummariseFillerClasses
FIXTURE  stock_01.xlsx sheet=Stock cols=A:Batch No,B:Filler Class,C:Qty On Hand,D:Reorder Level,E:Expiry Date,F:Remark rows=13 data=A:seq(101,1);B:values("E7018"x5," e7018 "x1,"ER70S-6"x4,"E308L"x1,""x1,"E6013"x1)+boundary("E7018","E7018","ER70S-6","ER70S-6"," e7018 ","E308L","","ER70S-6","E7018","E6013","E7018","E7018","ER70S-6");C:int(0,60)+boundary(10,11,3,20,9,40,2,"",0,4,50,5,1);D:int(1,20)+boundary(10,10,10,5,10,5,5,5,"",8,20,2,6);E:reldates(-60,120)+boundary(50,50,-1,0,"",-30,10,20,20,5,-10,-3,90);F:values("old"x7,"Reorder"x2,"Expired"x2,""x2)+boundary("old","Reorder","old","Expired","","old","old","Reorder","Expired","old","","old","old")
FIXTURE  stock_01.xlsx sheet=Reorder cols=A:Batch No,B:Filler Class,C:Qty On Hand,D:Reorder Level,E:Expiry Date,F:Remark rows=0 stale=7
FIXTURE  stock_01.xlsx sheet=Summary cols=A:Filler Class,B:Total Qty On Hand,C:Expired Batches rows=0 stale=6
FIXTURE  stock_02.xlsx sheet=Stock cols=A:Batch No,B:Filler Class,C:Qty On Hand,D:Reorder Level,E:Expiry Date,F:Remark rows=3 data=A:seq(201,1);B:values("E7018"x2,"E6013"x1)+boundary("E7018","E6013","E7018");C:int(1,60)+boundary(30,20,10);D:int(1,20)+boundary(5,5,5);E:reldates(1,90)+boundary(10,20,30);F:values("Expired"x1,"Reorder"x1,"old"x1)+boundary("Expired","Reorder","old")
FIXTURE  stock_02.xlsx sheet=Reorder cols=A:Batch No,B:Filler Class,C:Qty On Hand,D:Reorder Level,E:Expiry Date,F:Remark rows=0 stale=2
FIXTURE  stock_02.xlsx sheet=Summary cols=A:Filler Class,B:Total Qty On Hand,C:Expired Batches rows=0 stale=5
TEST     fixture=stock_01.xlsx call=modMain.RunHeadless(Stock, Reorder, Summary) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(Stock![Remark], "Expired") = 4 AND COUNTIF(Stock![Remark], "Reorder") = 5 AND COUNTA(Stock![Remark]) = 10 AND COUNTIF(Stock![Remark], "old") = 0 AND VALUE(Stock![Remark]2) = "Reorder" AND VALUE(Stock![Remark]4) = "Expired" AND VALUE(Stock![Remark]6) = "Reorder" AND VALUE(Stock![Remark]13) = "Expired" AND VALUE(Stock![Remark]14) = "Reorder" AND ROWS_USED(Reorder) = 6 AND COUNTIF(Reorder![Batch No], "stale") = 0 AND FILL_COUNT(Reorder![Batch No], #FF0000) = 0 AND VALUE(Reorder![Batch No]2) = 113 AND VALUE(Reorder![Batch No]3) = 107 AND VALUE(Reorder![Batch No]4) = 110 AND VALUE(Reorder![Batch No]5) = 105 AND VALUE(Reorder![Batch No]6) = 101 AND ROWS_USED(Summary) = 5 AND FILL_COUNT(Summary![Filler Class], #FF0000) = 0 AND VALUE(Summary![Filler Class]2) = "E308L" AND VALUE(Summary![Total Qty On Hand]2) = 0 AND VALUE(Summary![Expired Batches]2) = 1 AND VALUE(Summary![Filler Class]3) = "E6013" AND VALUE(Summary![Total Qty On Hand]3) = 4 AND VALUE(Summary![Expired Batches]3) = 0 AND VALUE(Summary![Filler Class]4) = "E7018" AND VALUE(Summary![Total Qty On Hand]4) = 30 AND VALUE(Summary![Expired Batches]4) = 2 AND VALUE(Summary![Filler Class]5) = "ER70S-6" AND VALUE(Summary![Total Qty On Hand]5) = 21 AND VALUE(Summary![Expired Batches]5) = 1
TEST     fixture=stock_01.xlsx call=modMain.RunHeadless(Stock, Reorder, Summary) assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(Stock![Batch No]) = PRE_COUNTA(Stock![Batch No]) AND SUM(Stock![Qty On Hand]) = PRE_SUM(Stock![Qty On Hand]) AND COUNTA(Stock![Expiry Date]) = PRE_COUNTA(Stock![Expiry Date]) AND VALUE(Reorder![Filler Class]2) = "ER70S-6" AND VALUE(Reorder![Qty On Hand]2) = 1 AND VALUE(Reorder![Reorder Level]2) = 6 AND VALUE(Reorder![Remark]2) = "Reorder" AND COUNTA(Reorder![Expiry Date]) = 5 AND COUNTIF(Reorder![Remark], "Reorder") = 5 AND SUM(Reorder![Qty On Hand]) = 26 AND SUM(Summary![Total Qty On Hand]) = 55 AND SUM(Summary![Expired Batches]) = 4
TEST     fixture=stock_01.xlsx call=modMain.RunHeadless(Summary, Reorder, Stock) assert=RESULT = -1 AND ERROR_COUNT = 1
TEST     fixture=stock_02.xlsx call=modMain.RunHeadless(Stock, Reorder, Summary) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTA(Stock![Remark]) = 1 AND ROWS_USED(Reorder) = 1 AND COUNTA(Reorder![Batch No]) = 1 AND ROWS_USED(Summary) = 3 AND VALUE(Summary![Filler Class]2) = "E6013" AND VALUE(Summary![Total Qty On Hand]2) = 20 AND VALUE(Summary![Expired Batches]2) = 0 AND VALUE(Summary![Filler Class]3) = "E7018" AND VALUE(Summary![Total Qty On Hand]3) = 40 AND VALUE(Summary![Expired Batches]3) = 0
---
VCL 0.6
TASK
GOAL     Start the consumables stock review on the workbook's own sheets and tell the user how many batches are expired.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "Consumable Stock"), SHEET(THIS_BOOK, "Reorder List"), SHEET(THIS_BOOK, "Class Summary"), into=n)
OP       IF n < 0
OP         MSG("The stock review did not run: check the headings of the Consumable Stock sheet, then look in ForgeLog.")
OP       ELSE
OP         MSG("Expired batches: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run the three steps of the consumables stock review on the sheets it is given and return how many batches are expired.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     stock As Worksheet, reorder As Worksheet, summary As Worksheet
RET      Long
VAR      expired As Long
PRE      COL($stock, "Batch No") > 0; COL($stock, "Filler Class") > 0; COL($stock, "Qty On Hand") > 0; COL($stock, "Reorder Level") > 0; COL($stock, "Expiry Date") > 0; COL($stock, "Remark") > 0
MODIFIES $stock["Remark"]
MODIFIES $reorder
MODIFIES $summary
OP       CALL(modSteps.MarkStockRemarks, $stock, into=expired)
OP       CALL(modSteps.ListReorderBatches, $stock, $reorder)
OP       CALL(modSteps.SummariseFillerClasses, $stock, $summary)
OP       RETURN(expired)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Expired in Remark for each batch past its Expiry Date, otherwise Reorder when Qty On Hand is at or below Reorder Level, clear the other remarks and return the number expired.
PROC     MarkStockRemarks
KIND     Function
MODULE   modSteps
ARGS     stock As Worksheet
RET      Long
VAR      cBatch As Long
VAR      cQty As Long
VAR      cLevel As Long
VAR      cExpiry As Long
VAR      cRemark As Long
VAR      remark As String
VAR      expired As Long
PRE      COL($stock, "Batch No") > 0; COL($stock, "Qty On Hand") > 0; COL($stock, "Reorder Level") > 0; COL($stock, "Expiry Date") > 0; COL($stock, "Remark") > 0
MODIFIES $stock["Remark"]
OP       LET(cBatch, COL($stock, "Batch No"))
OP       LET(cQty, COL($stock, "Qty On Hand"))
OP       LET(cLevel, COL($stock, "Reorder Level"))
OP       LET(cExpiry, COL($stock, "Expiry Date"))
OP       LET(cRemark, COL($stock, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($stock, cBatch) var=r
OP         LET(remark, "")
OP         IF IS_NUMBER(CELL($stock, r, cQty))
OP           IF IS_NUMBER(CELL($stock, r, cLevel))
OP             IF CELL($stock, r, cQty) <= CELL($stock, r, cLevel) THEN LET(remark, "Reorder")
OP         IF IS_DATE(CELL($stock, r, cExpiry))
OP           IF CELL($stock, r, cExpiry) < TODAY
OP             LET(remark, "Expired")
OP             LET(expired, expired + 1)
OP         IF remark = ""
OP           CLEAR(CELL($stock, r, cRemark), what=contents)
OP         ELSE
OP           SET_VALUE(CELL($stock, r, cRemark), remark)
OP       RETURN(expired)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Reorder List sheet with every batch whose Remark is Reorder, under the stock sheet's six headings, the lowest Qty On Hand first.
PROC     ListReorderBatches
KIND     Sub
MODULE   modSteps
ARGS     stock As Worksheet, reorder As Worksheet
VAR      cBatch As Long
VAR      cClass As Long
VAR      cQty As Long
VAR      cLevel As Long
VAR      cExpiry As Long
VAR      cRemark As Long
VAR      outRow As Long
PRE      COL($stock, "Batch No") > 0; COL($stock, "Filler Class") > 0; COL($stock, "Qty On Hand") > 0; COL($stock, "Reorder Level") > 0; COL($stock, "Expiry Date") > 0; COL($stock, "Remark") > 0
MODIFIES $reorder
OP       LET(cBatch, COL($stock, "Batch No"))
OP       LET(cClass, COL($stock, "Filler Class"))
OP       LET(cQty, COL($stock, "Qty On Hand"))
OP       LET(cLevel, COL($stock, "Reorder Level"))
OP       LET(cExpiry, COL($stock, "Expiry Date"))
OP       LET(cRemark, COL($stock, "Remark"))
OP       CLEAR(USED_RANGE($reorder), what=all)
OP       WRITE_ROW($reorder, 1, "Batch No", "Filler Class", "Qty On Hand", "Reorder Level", "Expiry Date", "Remark")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($stock, cBatch) var=r
OP         IF SAME(CELL($stock, r, cRemark), "Reorder")
OP           WRITE_ROW($reorder, outRow, CELL($stock, r, cBatch), CELL($stock, r, cClass), CELL($stock, r, cQty), CELL($stock, r, cLevel), CELL($stock, r, cExpiry), CELL($stock, r, cRemark))
OP           LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($reorder), key=3, order=asc, header=yes)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Class Summary sheet with one row per Filler Class, A to Z, showing its total Qty On Hand without expired batches and its number of expired batches.
PROC     SummariseFillerClasses
KIND     Sub
MODULE   modSteps
ARGS     stock As Worksheet, summary As Worksheet
VAR      cBatch As Long
VAR      cClass As Long
VAR      cQty As Long
VAR      cRemark As Long
VAR      cls As String
VAR      totals As Dictionary
VAR      expiredCounts As Dictionary
VAR      total As Double
VAR      nExpired As Long
VAR      outRow As Long
PRE      COL($stock, "Batch No") > 0; COL($stock, "Filler Class") > 0; COL($stock, "Qty On Hand") > 0; COL($stock, "Remark") > 0
MODIFIES $summary
OP       LET(cBatch, COL($stock, "Batch No"))
OP       LET(cClass, COL($stock, "Filler Class"))
OP       LET(cQty, COL($stock, "Qty On Hand"))
OP       LET(cRemark, COL($stock, "Remark"))
OP       CLEAR(USED_RANGE($summary), what=all)
OP       WRITE_ROW($summary, 1, "Filler Class", "Total Qty On Hand", "Expired Batches")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($stock, cBatch) var=r
OP         SKIP_IF IS_EMPTY(CELL($stock, r, cClass))
OP         LET(cls, TRIM(CELL($stock, r, cClass)))
OP         DICT_ADD(totals, cls, 0, mode=sum)
OP         IF SAME(CELL($stock, r, cRemark), "Expired")
OP           DICT_ADD(expiredCounts, cls, 1, mode=sum)
OP         ELSE
OP           IF IS_NUMBER(CELL($stock, r, cQty)) THEN DICT_ADD(totals, cls, CELL($stock, r, cQty), mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(totals)
OP         DICT_GET(totals, k, into=total)
OP         LET(nExpired, 0)
OP         IF EXISTS(expiredCounts, k) THEN DICT_GET(expiredCounts, k, into=nExpired)
OP         WRITE_ROW($summary, outRow, k, total, nExpired)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($summary), key=1, order=asc, header=yes)
ERR      standard
