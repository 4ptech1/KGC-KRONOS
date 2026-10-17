VCL 0.1
TASK
# ASSUME: Consumable Issue headers (row 1) include "Batch No" and "Qty Issued"
# ASSUME: Inventory headers include "Batch No", "Opening Stock", "Issued to Date", "Balance" and "Remark"
# ASSUME: Issued to Date, Balance and Remark are owned by this macro and recalculated from the whole issue log on every run, so running it twice never double-counts
# ASSUME: issue rows with a blank or non-numeric Qty Issued, or a blank Batch No, are ignored
# ASSUME: the function returns how many issue rows use a batch number that is not in Inventory
PROC     ReconcileConsumableStock
KIND     Function
MODULE   modWelding
ARGS     issues As Worksheet, inv As Worksheet
RET      Long
VAR      known As Dictionary
VAR      issued As Dictionary
VAR      iBatch As Long
VAR      iQty As Long
VAR      vBatch As Long
VAR      vOpen As Long
VAR      vIssued As Long
VAR      vBal As Long
VAR      vRemark As Long
VAR      batch As String
VAR      used As Double
VAR      unknown As Long
PRE      COL($issues, "Batch No") > 0; COL($issues, "Qty Issued") > 0; COL($inv, "Batch No") > 0; COL($inv, "Opening Stock") > 0; COL($inv, "Issued to Date") > 0; COL($inv, "Balance") > 0; COL($inv, "Remark") > 0
OP       LET(iBatch, COL($issues, "Batch No"))
OP       LET(iQty, COL($issues, "Qty Issued"))
OP       LET(vBatch, COL($inv, "Batch No"))
OP       LET(vOpen, COL($inv, "Opening Stock"))
OP       LET(vIssued, COL($inv, "Issued to Date"))
OP       LET(vBal, COL($inv, "Balance"))
OP       LET(vRemark, COL($inv, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($inv, vBatch) var=r
OP         IF NOT IS_EMPTY(CELL($inv, r, vBatch)) THEN DICT_ADD(known, TRIM(CELL($inv, r, vBatch)), r, mode=keep)
OP       LET(unknown, 0)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($issues, iBatch) var=r
OP         LET(batch, TRIM(CELL($issues, r, iBatch)))
OP         IF NOT IS_EMPTY(batch) AND IS_NUMBER(CELL($issues, r, iQty))
OP           IF EXISTS(known, batch) THEN DICT_ADD(issued, batch, CELL($issues, r, iQty), mode=sum)
OP           IF NOT EXISTS(known, batch) THEN LET(unknown, unknown + 1)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($inv, vBatch) var=r
OP         IF NOT IS_EMPTY(CELL($inv, r, vBatch))
OP           LET(batch, TRIM(CELL($inv, r, vBatch)))
OP           LET(used, 0)
OP           IF EXISTS(issued, batch) THEN DICT_GET(issued, batch, into=used)
OP           SET_VALUE(CELL($inv, r, vIssued), used)
OP           CLEAR(CELL($inv, r, vRemark), what=contents)
OP           IF IS_NUMBER(CELL($inv, r, vOpen))
OP             SET_VALUE(CELL($inv, r, vBal), CELL($inv, r, vOpen) - used)
OP             IF CELL($inv, r, vOpen) - used < 0 THEN SET_VALUE(CELL($inv, r, vRemark), "Negative balance")
OP           ELSE
OP             CLEAR(CELL($inv, r, vBal), what=contents)
OP             SET_VALUE(CELL($inv, r, vRemark), "Opening stock missing")
OP       RETURN(unknown)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  consumables_01.xlsx sheet=Inventory cols=A:Batch No,B:Opening Stock,C:Issued to Date,D:Balance,E:Remark rows=9 data=A:values("B-101"x1,"B-102"x1,"B-103"x1,"B-104"x1,"B-105"x1,"B-106"x1,"B-107"x1,"B-108"x1,""x1)+boundary("B-101","B-106");B:int(200,500)+boundary(1,0)+blank(1);D:int(1,99);E:values("Old remark"x5,""x4)
FIXTURE  consumables_01.xlsx sheet=Issues cols=A:Issue Date,B:Batch No,C:Qty Issued rows=60 data=A:dates(2026-06-01,2026-09-30);B:values("B-101"x10,"B-102"x10,"B-103"x10,"B-104"x10,"B-105"x10,"B-109"x5,""x5);C:int(1,20)+blank(6)
TEST     fixture=consumables_01.xlsx call=ReconcileConsumableStock(Issues, Inventory) runs=2 assert=RESULT = PRE_COUNTIFS(Issues![Batch No], "B-109", Issues![Qty Issued], "<>") AND SUM(Inventory![Issued to Date]) = PRE_SUMIFS(Issues![Qty Issued], Issues![Batch No], "<>B-109", Issues![Batch No], "<>") AND ERROR_COUNT = 0
TEST     fixture=consumables_01.xlsx call=ReconcileConsumableStock(Issues, Inventory) assert=COUNTIF(Inventory![Remark], "Opening stock missing") = 1 AND COUNTIF(Inventory![Remark], "Negative balance") = 1 AND COUNTIF(Inventory![Remark], "Old remark") = 1 AND VALUE(Inventory![Balance]3) = 0 AND COUNTA(Inventory![Balance]) = 9
TEST     fixture=consumables_01.xlsx call=ReconcileConsumableStock(Issues, Inventory) assert=SUM(Inventory![Opening Stock]) = PRE_SUM(Inventory![Opening Stock]) AND SUM(Issues![Qty Issued]) = PRE_SUM(Issues![Qty Issued])
TEST     fixture=consumables_01.xlsx call=ReconcileConsumableStock(Inventory, Inventory) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Inventory![Balance]) = PRE_COUNTA(Inventory![Balance])
