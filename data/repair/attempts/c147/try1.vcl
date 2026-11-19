VCL 0.6
BUILD
# ASSUME: Weld Log headers (row 1) are "Weld Seam No", "Weld Date", "Batch No", "Remark"; Consumable Stock has "Batch No", "Filler Class", "Expiry Date"
# ASSUME: batch numbers are matched after trimming and ignoring case; an empty Batch No in the Weld Log counts as not in stock (Unknown batch)
# ASSUME: a batch is Expired only when its Expiry Date is a real date before a real Weld Date; a missing or text date gives no remark
# ASSUME: the output sheets Batch Usage and Weld Problems already exist and are rebuilt on every run; Batch Usage has the headings Batch No and Welds
# ASSUME: the real sheets are "Weld Log", "Consumable Stock", "Batch Usage" and "Weld Problems"; the TESTs use WeldLog, Stock, Usage and Problems
# ASSUME: step 3 lists every Weld Log row whose Remark is not blank, with the Weld Log headings; step 4 counts the same rows and shows a message
GOAL     Mark each weld with an unknown or expired filler batch, then list batch usage and the problem welds and count the welds with a remark.
NAME     FillerTraceability
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunMarkRemarks()
SIG      modMain.RunBatchUsage()
SIG      modMain.RunProblems()
SIG      modMain.RunCountRemarks()
SIG      modMain.RunHeadless(log As Worksheet, stock As Worksheet, usage As Worksheet, problems As Worksheet) As Long
SIG      modSteps.MarkRemarks(log As Worksheet, stock As Worksheet) As Long
SIG      modSteps.WriteBatchUsage(log As Worksheet, usage As Worksheet)
SIG      modSteps.WriteProblems(log As Worksheet, problems As Worksheet)
SIG      modSteps.CountRemarks(log As Worksheet) As Long
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.MarkRemarks
FLOW     modMain.RunHeadless -> modSteps.WriteBatchUsage
FLOW     modMain.RunHeadless -> modSteps.WriteProblems
FLOW     modMain.RunHeadless -> modSteps.CountRemarks
FLOW     modMain.RunMarkRemarks -> modSteps.MarkRemarks
FLOW     modMain.RunBatchUsage -> modSteps.WriteBatchUsage
FLOW     modMain.RunProblems -> modSteps.WriteProblems
FLOW     modMain.RunCountRemarks -> modSteps.CountRemarks
FIXTURE  fill_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Weld Date,C:Batch No,D:Remark rows=9 data=A:seq(1,1);B:reldates(-30,0)+boundary(-10,-9,-10,-10,-3,-3,-3,"",-20);C:values("B-100"x1,"  b-100 "x1,"B-200"x1,"B-300"x2,"B-999"x1,""x1,"B-400"x1,"b-100"x1)+boundary("B-100","  b-100 ","B-200","B-300","B-999","","B-400","B-300","b-100");D:values("old"x9)
FIXTURE  fill_01.xlsx sheet=Stock cols=A:Batch No,B:Filler Class,C:Expiry Date rows=6 data=A:values("B-100"x1,"b-200"x1,"B-300"x1,"B-400"x1,""x1,"B-500"x1)+boundary("B-100","b-200","B-300","B-400","","B-500");B:text(6);C:reldates(-30,30)+boundary(-10,-11,5,"",3,-1)
FIXTURE  fill_01.xlsx sheet=Usage cols=A:Batch No,B:Welds rows=0 stale=2
FIXTURE  fill_01.xlsx sheet=Problems cols=A:Weld Seam No,B:Weld Date,C:Batch No,D:Remark rows=0 stale=2
TEST     fixture=fill_01.xlsx call=modMain.RunHeadless(WeldLog, Stock, Usage, Problems) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(WeldLog![Remark], "Unknown batch") = 2 AND COUNTIF(WeldLog![Remark], "Expired batch") = 2 AND COUNTIF(WeldLog![Remark], "old") = 0 AND COUNTA(Usage![Batch No]) = 6 AND VALUE(Usage![Welds]2) = 3 AND VALUE(Usage![Welds]3) = 2 AND COUNTIF(Usage![Batch No], "stale") = 0 AND COUNTA(Problems![Weld Seam No]) = 5 AND COUNTIF(Problems![Remark], "stale") = 0
TEST     fixture=fill_01.xlsx call=modMain.RunHeadless(WeldLog, Stock, Usage, Problems) assert=RESULT = 4 AND COUNTA(WeldLog![Weld Seam No]) = PRE_COUNTA(WeldLog![Weld Seam No]) AND COUNTA(Stock![Batch No]) = PRE_COUNTA(Stock![Batch No]) AND VALUE(Usage![Batch No]2) = "B-100" AND VALUE(Problems![Remark]2) = "Expired batch" AND VALUE(Problems![Weld Seam No]2) = 2 AND VALUE(Problems![Weld Seam No]3) = 3 AND VALUE(Problems![Remark]4) = "Unknown batch"
TEST     fixture=fill_01.xlsx call=modMain.RunHeadless(Stock, WeldLog, Usage, Problems) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
TASK
GOAL     Start the whole filler traceability tool on the workbook's own sheets and tell the user how many welds have a remark.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "Weld Log"), SHEET(THIS_BOOK, "Consumable Stock"), SHEET(THIS_BOOK, "Batch Usage"), SHEET(THIS_BOOK, "Weld Problems"), into=n)
OP       MSG("Welds with a remark: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run step 1 on its own: write the batch remarks in the Weld Log and tell the user how many welds got one.
PROC     RunMarkRemarks
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modSteps.MarkRemarks, SHEET(THIS_BOOK, "Weld Log"), SHEET(THIS_BOOK, "Consumable Stock"), into=n)
OP       MSG("Welds given a remark: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run step 2 on its own: rebuild the Batch Usage sheet from the Weld Log.
PROC     RunBatchUsage
KIND     Sub
MODULE   modMain
OP       CALL(modSteps.WriteBatchUsage, SHEET(THIS_BOOK, "Weld Log"), SHEET(THIS_BOOK, "Batch Usage"))
ERR      standard
---
VCL 0.6
TASK
GOAL     Run step 3 on its own: rebuild the Weld Problems sheet with the welds that have a remark.
PROC     RunProblems
KIND     Sub
MODULE   modMain
OP       CALL(modSteps.WriteProblems, SHEET(THIS_BOOK, "Weld Log"), SHEET(THIS_BOOK, "Weld Problems"))
ERR      standard
---
VCL 0.6
TASK
GOAL     Run step 4 on its own: tell the user how many welds in the Weld Log have a remark.
PROC     RunCountRemarks
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modSteps.CountRemarks, SHEET(THIS_BOOK, "Weld Log"), into=n)
OP       MSG("Welds with a remark: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run all four steps on the four sheets it is given and return how many welds have a remark.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     log As Worksheet, stock As Worksheet, usage As Worksheet, problems As Worksheet
RET      Long
VAR      total As Long
PRE      COL($log, "Weld Seam No") > 0; COL($log, "Weld Date") > 0; COL($log, "Batch No") > 0; COL($log, "Remark") > 0; COL($stock, "Batch No") > 0; COL($stock, "Expiry Date") > 0
MODIFIES $log["Remark"]
MODIFIES $usage
MODIFIES $problems
OP       CALL(modSteps.MarkRemarks, $log, $stock)
OP       CALL(modSteps.WriteBatchUsage, $log, $usage)
OP       CALL(modSteps.WriteProblems, $log, $problems)
OP       CALL(modSteps.CountRemarks, $log, into=total)
OP       RETURN(total)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Unknown batch or Expired batch in the Weld Log Remark column, clear it on the other welds and return how many got a remark.
PROC     MarkRemarks
KIND     Function
MODULE   modSteps
ARGS     log As Worksheet, stock As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cDate As Long
VAR      cBatch As Long
VAR      cRemark As Long
VAR      cStockBatch As Long
VAR      cExpiry As Long
VAR      expiries As Dictionary
VAR      batch As String
VAR      remark As String
VAR      expiryDate As Variant
VAR      n As Long
PRE      COL($log, "Weld Seam No") > 0; COL($log, "Weld Date") > 0; COL($log, "Batch No") > 0; COL($log, "Remark") > 0; COL($stock, "Batch No") > 0; COL($stock, "Expiry Date") > 0
MODIFIES $log["Remark"]
OP       LET(cSeam, COL($log, "Weld Seam No"))
OP       LET(cDate, COL($log, "Weld Date"))
OP       LET(cBatch, COL($log, "Batch No"))
OP       LET(cRemark, COL($log, "Remark"))
OP       LET(cStockBatch, COL($stock, "Batch No"))
OP       LET(cExpiry, COL($stock, "Expiry Date"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($stock, cStockBatch) var=s
OP         SKIP_IF IS_EMPTY(CELL($stock, s, cStockBatch))
OP         DICT_ADD(expiries, TRIM(CELL($stock, s, cStockBatch)), CELL($stock, s, cExpiry), mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($log, cSeam) var=r
OP         LET(batch, TRIM(CELL($log, r, cBatch)))
OP         LET(remark, "")
OP         IF EXISTS(expiries, batch)
OP           DICT_GET(expiries, batch, into=expiryDate)
OP           IF IS_DATE(expiryDate)
OP             IF IS_DATE(CELL($log, r, cDate))
OP               IF expiryDate < CELL($log, r, cDate) THEN LET(remark, "Expired batch")
OP         ELSE
OP           LET(remark, "Unknown batch")
OP         IF remark = ""
OP           CLEAR(CELL($log, r, cRemark), what=contents)
OP         ELSE
OP           SET_VALUE(CELL($log, r, cRemark), remark)
OP           LET(n, n + 1)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Batch Usage sheet with each batch number used in the Weld Log and its number of welds, the most welds first.
PROC     WriteBatchUsage
KIND     Sub
MODULE   modSteps
ARGS     log As Worksheet, usage As Worksheet
VAR      cBatch As Long
VAR      counts As Dictionary
VAR      got As Long
VAR      outRow As Long
PRE      COL($log, "Batch No") > 0
MODIFIES $usage
OP       LET(cBatch, COL($log, "Batch No"))
OP       CLEAR(USED_RANGE($usage), what=all)
OP       WRITE_ROW($usage, 1, "Batch No", "Welds")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($log, cBatch) var=r
OP         SKIP_IF IS_EMPTY(CELL($log, r, cBatch))
OP         DICT_ADD(counts, TRIM(CELL($log, r, cBatch)), 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         WRITE_ROW($usage, outRow, k, got)
OP         LET(outRow, outRow + 1)
OP       IF outRow > 3 THEN SORT(USED_RANGE($usage), key=2, order=desc, header=yes)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Weld Problems sheet with the Weld Log headings and every weld that has a remark.
PROC     WriteProblems
KIND     Sub
MODULE   modSteps
ARGS     log As Worksheet, problems As Worksheet
VAR      cSeam As Long
VAR      cRemark As Long
VAR      nCols As Long
VAR      outRow As Long
PRE      COL($log, "Weld Seam No") > 0; COL($log, "Remark") > 0
MODIFIES $problems
OP       LET(cSeam, COL($log, "Weld Seam No"))
OP       LET(cRemark, COL($log, "Remark"))
OP       LET(nCols, LASTCOL($log, 1))
OP       CLEAR(USED_RANGE($problems), what=all)
OP       COPY_ROW($log, 1, $problems, 1, cols=nCols)
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($log, cSeam) var=r
OP         IF NOT IS_EMPTY(CELL($log, r, cRemark))
OP           COPY_ROW($log, r, $problems, outRow, cols=nCols)
OP           LET(outRow, outRow + 1)
ERR      standard
---
VCL 0.6
TASK
GOAL     Count the welds in the Weld Log that have a remark and return that number.
PROC     CountRemarks
KIND     Function
MODULE   modSteps
ARGS     log As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cRemark As Long
VAR      n As Long
PRE      COL($log, "Weld Seam No") > 0; COL($log, "Remark") > 0
OP       LET(cSeam, COL($log, "Weld Seam No"))
OP       LET(cRemark, COL($log, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($log, cSeam) var=r
OP         IF NOT IS_EMPTY(CELL($log, r, cRemark)) THEN LET(n, n + 1)
OP       RETURN(n)
ERR      standard
