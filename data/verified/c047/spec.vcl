VCL 0.6
TASK
# ASSUME: Stock headers in row 1 include "Batch No", "Qty On Hand", "Unit Cost" and "Value"
# ASSUME: Qty On Hand is in column B and Unit Cost in column C, because the formula text names those columns; if not, nothing changes
# ASSUME: data rows run from row 2 to the last row that has a Batch No
# ASSUME: every such row gets the formula, whether its Qty On Hand and Unit Cost are filled or not; an old value in Value is replaced
# ASSUME: the Value column is not formatted as Text
GOAL     Put a formula in the Value column of every Stock data row that multiplies Qty On Hand by Unit Cost, and return how many formulas were written.
PROC     WriteStockValues
KIND     Function
MODULE   modStock
ARGS     ws As Worksheet
RET      Long
VAR      cBatch As Long
VAR      cValue As Long
VAR      written As Long
PRE      COL($ws, "Batch No") > 0; COL($ws, "Qty On Hand") = 2; COL($ws, "Unit Cost") = 3; COL($ws, "Value") > 0
MODIFIES $ws["Value"]
OP       LET(cBatch, COL($ws, "Batch No"))
OP       LET(cValue, COL($ws, "Value"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cBatch) var=r
OP         SET_FORMULA(CELL($ws, r, cValue), "=RC2*RC3", style=r1c1)
OP         LET(written, written + 1)
OP       RETURN(written)
ERR      standard
FIXTURE  stock_01.xlsx sheet=Stock cols=A:Batch No,B:Qty On Hand,C:Unit Cost,D:Value rows=6 data=A:seq(1,1);B:int(0,100)+boundary(0,5);C:num(1,50)+boundary(2.5);D:values("old"x6)
FIXTURE  stock_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  stock_02.xlsx sheet=Stock cols=A:Batch No,B:Unit Cost,C:Qty On Hand,D:Value rows=3 data=A:seq(1,1);B:num(1,50);C:int(0,100);D:values("old"x3)
TEST     fixture=stock_01.xlsx call=WriteStockValues(Stock) runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND FORMULA(Stock![Value]2) = "=RC2*RC3" AND FORMULA(Stock![Value]3) = "=RC2*RC3" AND FORMULA(Stock![Value]4) = "=RC2*RC3" AND FORMULA(Stock![Value]5) = "=RC2*RC3" AND FORMULA(Stock![Value]6) = "=RC2*RC3" AND FORMULA(Stock![Value]7) = "=RC2*RC3"
TEST     fixture=stock_01.xlsx call=WriteStockValues(Stock) assert=FORMULA(Stock![Value]1) = "" AND COUNTA(Stock![Batch No]) = PRE_COUNTA(Stock![Batch No]) AND SUM(Stock![Qty On Hand]) = PRE_SUM(Stock![Qty On Hand]) AND SUM(Stock![Unit Cost]) = PRE_SUM(Stock![Unit Cost])
TEST     fixture=stock_01.xlsx call=WriteStockValues(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
TEST     fixture=stock_02.xlsx call=WriteStockValues(Stock) assert=RESULT = -1 AND ERROR_COUNT = 1 AND FORMULA(Stock![Value]2) = ""
