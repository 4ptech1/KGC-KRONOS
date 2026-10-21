VCL 0.6
TASK
# ASSUME: Consumable Stock headers in row 1 include "Batch No", "Qty On Hand", "Reorder Level" and "Remark"
# ASSUME: the Remark column is used only for this flag, so a Remark is cleared whenever Qty On Hand is above Reorder Level
# ASSUME: a row where Qty On Hand or Reorder Level is not a number keeps its Remark
# ASSUME: a row with a blank Batch No is skipped
GOAL     Put Reorder in Remark for each batch at or below its reorder level, clear the Remark of each batch above it, and return how many batches need reordering.
PROC     FlagBatchesToReorder
KIND     Function
MODULE   modStock
ARGS     ws As Worksheet
RET      Long
VAR      cBatch As Long
VAR      cQty As Long
VAR      cLevel As Long
VAR      cRemark As Long
VAR      due As Long
PRE      COL($ws, "Batch No") > 0; COL($ws, "Qty On Hand") > 0; COL($ws, "Reorder Level") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Remark"]
OP       LET(cBatch, COL($ws, "Batch No"))
OP       LET(cQty, COL($ws, "Qty On Hand"))
OP       LET(cLevel, COL($ws, "Reorder Level"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cBatch) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cBatch))
OP         IF IS_NUMBER(CELL($ws, r, cQty))
OP           IF IS_NUMBER(CELL($ws, r, cLevel))
OP             IF CELL($ws, r, cQty) <= CELL($ws, r, cLevel)
OP               SET_VALUE(CELL($ws, r, cRemark), "Reorder")
OP               LET(due, due + 1)
OP             ELSE
OP               CLEAR(CELL($ws, r, cRemark), what=contents)
OP       RETURN(due)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  stock_01.xlsx sheet=ConsumableStock cols=A:Batch No,B:Description,C:Qty On Hand,D:Reorder Level,E:Remark rows=9 data=A:text(5)+boundary("B-01","B-02","B-03","B-04","B-05","","B-06","B-07","B-08");B:text(10);C:int(0,100)+boundary(5,10,11,0,20,1,"",3,50);D:int(5,10)+boundary(10,10,10,5,5,10,10,"",10);E:values("note"x6,"Reorder"x3)+boundary("note","note","Reorder","note","Reorder","note","note","note","Reorder")
FIXTURE  stock_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=stock_01.xlsx call=FlagBatchesToReorder(ConsumableStock) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(ConsumableStock![Remark], "Reorder") = 3 AND COUNTIF(ConsumableStock![Remark], "note") = 3 AND COUNTA(ConsumableStock![Remark]) = 7
TEST     fixture=stock_01.xlsx call=FlagBatchesToReorder(ConsumableStock) assert=RESULT = 3 AND COUNTA(ConsumableStock![Batch No]) = PRE_COUNTA(ConsumableStock![Batch No]) AND SUM(ConsumableStock![Qty On Hand]) = PRE_SUM(ConsumableStock![Qty On Hand]) AND SUM(ConsumableStock![Reorder Level]) = PRE_SUM(ConsumableStock![Reorder Level]) AND VALUE(ConsumableStock![Remark]2) = "Reorder" AND VALUE(ConsumableStock![Remark]3) = "Reorder" AND COUNTIF(ConsumableStock![Remark]4, "Reorder") = 0 AND VALUE(ConsumableStock![Remark]5) = "Reorder" AND COUNTIF(ConsumableStock![Remark]6, "Reorder") = 0 AND VALUE(ConsumableStock![Remark]7) = "note" AND VALUE(ConsumableStock![Remark]8) = "note" AND VALUE(ConsumableStock![Remark]9) = "note" AND COUNTIF(ConsumableStock![Remark]10, "Reorder") = 0
TEST     fixture=stock_01.xlsx call=FlagBatchesToReorder(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
