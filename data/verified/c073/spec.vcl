VCL 0.6
TASK
# ASSUME: the sheet's headers in row 1 include "Batch No", "Expiry Date" and "Remark"; every batch has a Batch No, so rows are read down to the last Batch No
# ASSUME: a batch that expires today is not yet expired: it gets "Expiring soon", as does every date from today to today + 30 days
# ASSUME: an Expiry Date that is blank or is not a real Excel date (typed text) counts as no date: its Remark is cleared
# ASSUME: the count of expired batches is returned and shown in a message
GOAL     Write Expired or Expiring soon in the Remark column of the Consumable Stock sheet from each batch's Expiry Date, clear the other remarks and count the expired batches.
PROC     FlagExpiringBatches
KIND     Function
MODULE   modConsumables
ARGS     ws As Worksheet
RET      Long
VAR      cBatch As Long
VAR      cExpiry As Long
VAR      cRemark As Long
VAR      expired As Long
PRE      COL($ws, "Batch No") > 0; COL($ws, "Expiry Date") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Remark"]
OP       LET(cBatch, COL($ws, "Batch No"))
OP       LET(cExpiry, COL($ws, "Expiry Date"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LET(expired, 0)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cBatch) var=r
OP         IF IS_DATE(CELL($ws, r, cExpiry))
OP           IF CELL($ws, r, cExpiry) < TODAY
OP             SET_VALUE(CELL($ws, r, cRemark), "Expired")
OP             LET(expired, expired + 1)
OP           ELSE
OP             IF CELL($ws, r, cExpiry) <= TODAY + 30
OP               SET_VALUE(CELL($ws, r, cRemark), "Expiring soon")
OP             ELSE
OP               CLEAR(CELL($ws, r, cRemark), what=contents)
OP         ELSE
OP           CLEAR(CELL($ws, r, cRemark), what=contents)
OP       MSG("Expired batches: " & expired)
OP       RETURN(expired)
ERR      standard
FIXTURE  stock_01.xlsx sheet=Stock cols=A:Batch No,B:Description,C:Expiry Date,D:Qty On Hand,E:Remark rows=10 data=A:seq(1001,1);B:text(10);C:reldates(-60,90)+boundary(-1,0,30,31,-20,10,60,1,-5,"");D:int(1,50);E:values("Expired"x3,"Expiring soon"x2,"old"x4,""x1)+boundary("old","Expired","old","Expiring soon","Expiring soon","Expired","Expired","","old","old")
FIXTURE  stock_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=stock_01.xlsx call=FlagExpiringBatches(Stock) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Stock![Remark], "Expired") = 3 AND COUNTIF(Stock![Remark], "Expiring soon") = 4 AND COUNTIF(Stock![Remark], "old") = 0 AND COUNTA(Stock![Remark]) = 8 AND VALUE(Stock![Remark]2) = "Expired" AND VALUE(Stock![Remark]3) = "Expiring soon" AND VALUE(Stock![Remark]4) = "Expiring soon" AND VALUE(Stock![Remark]6) = "Expired" AND VALUE(Stock![Remark]10) = "Expired"
TEST     fixture=stock_01.xlsx call=FlagExpiringBatches(Stock) assert=RESULT = 3 AND COUNTA(Stock![Batch No]) = PRE_COUNTA(Stock![Batch No]) AND COUNTA(Stock![Expiry Date]) = PRE_COUNTA(Stock![Expiry Date]) AND SUM(Stock![Qty On Hand]) = PRE_SUM(Stock![Qty On Hand])
TEST     fixture=stock_01.xlsx call=FlagExpiringBatches(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
