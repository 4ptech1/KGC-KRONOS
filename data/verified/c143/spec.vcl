VCL 0.6
TASK
# ASSUME: Consumable Stock has headers in row 1: "Batch No", "Filler Class", "Qty On Hand", "Expiry Date"; the real sheets are "Consumable Stock" and "Expired Batches"
# ASSUME: the Expired Batches sheet already exists and is passed in; it is cleared and rebuilt on every run
# ASSUME: a batch is expired when its Expiry Date is a real date before today; a batch expiring today is not expired yet
# ASSUME: a blank or non-date Expiry Date is skipped; the oldest expiry comes first
# ASSUME: Consumable Stock is only read, and the four columns are copied in the order Batch No, Filler Class, Qty On Hand, Expiry Date
GOAL     Rebuild the Expired Batches sheet with every consumable batch whose expiry date is before today, oldest first, and return how many were listed.
PROC     ListExpiredBatches
KIND     Function
MODULE   modStock
ARGS     stock As Worksheet, out As Worksheet
RET      Long
VAR      cBatch As Long
VAR      cClass As Long
VAR      cQty As Long
VAR      cExp As Long
VAR      outRow As Long
PRE      COL($stock, "Batch No") > 0; COL($stock, "Filler Class") > 0; COL($stock, "Qty On Hand") > 0; COL($stock, "Expiry Date") > 0
MODIFIES $out
OP       LET(cBatch, COL($stock, "Batch No"))
OP       LET(cClass, COL($stock, "Filler Class"))
OP       LET(cQty, COL($stock, "Qty On Hand"))
OP       LET(cExp, COL($stock, "Expiry Date"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Batch No", "Filler Class", "Qty On Hand", "Expiry Date")
OP       LET(outRow, 1)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($stock, cExp) var=r
OP         IF IS_DATE(CELL($stock, r, cExp))
OP           IF CELL($stock, r, cExp) < TODAY
OP             LET(outRow, outRow + 1)
OP             WRITE_ROW($out, outRow, CELL($stock, r, cBatch), CELL($stock, r, cClass), CELL($stock, r, cQty), CELL($stock, r, cExp))
OP       IF outRow > 2
OP         SORT(USED_RANGE($out), key=4, order=asc, header=yes)
OP       RETURN(outRow - 1)
PERF     screen_off
ERR      standard
FIXTURE  stock_01.xlsx sheet=Stock cols=A:Batch No,B:Filler Class,C:Qty On Hand,D:Expiry Date rows=9 stale=2 data=A:seq(1,1);B:values("E7018"x5,"ER70S"x4);C:int(1,50);D:reldates(-30,5)+boundary(-1,0,1,"",-5,-30,5,-2,-10)
FIXTURE  stock_01.xlsx sheet=Expired cols=A:Batch No,B:Filler Class,C:Qty On Hand,D:Expiry Date rows=0 stale=3
FIXTURE  stock_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  stock_02.xlsx sheet=Stock cols=A:Batch No,B:Filler Class,C:Qty On Hand,D:Expiry Date rows=3 data=A:seq(1,1);B:values("E7018"x3);C:int(1,50);D:reldates(0,9)+boundary(0,1,"")
FIXTURE  stock_02.xlsx sheet=Expired cols=A:Batch No,B:Filler Class,C:Qty On Hand,D:Expiry Date rows=0 stale=2
TEST     fixture=stock_01.xlsx call=ListExpiredBatches(Stock, Expired) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND ROWS_USED(Expired) = 6 AND COUNTIF(Expired![Batch No], "stale") = 0 AND VALUE(Expired![Batch No]2) = 6 AND VALUE(Expired![Batch No]3) = 9 AND VALUE(Expired![Batch No]4) = 5 AND VALUE(Expired![Batch No]5) = 8 AND VALUE(Expired![Batch No]6) = 1
TEST     fixture=stock_01.xlsx call=ListExpiredBatches(Stock, Expired) assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTA(Stock![Batch No]) = PRE_COUNTA(Stock![Batch No]) AND COUNTA(Stock![Expiry Date]) = PRE_COUNTA(Stock![Expiry Date]) AND COUNTA(Expired![Filler Class]) = 6
TEST     fixture=stock_02.xlsx call=ListExpiredBatches(Stock, Expired) assert=RESULT = 0 AND ERROR_COUNT = 0 AND ROWS_USED(Expired) = 1 AND COUNTA(Expired![Batch No]) = 1
TEST     fixture=stock_01.xlsx call=ListExpiredBatches(Notes, Expired) assert=RESULT = -1 AND ERROR_COUNT = 1
