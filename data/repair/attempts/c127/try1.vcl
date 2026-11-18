VCL 0.6
TASK
# ASSUME: Consumable Stock headers are Batch No, Qty On Hand, Reorder Level and Reorder in columns A to D (the TEST sheet is called Stock); PRE checks this because a formula's text can't use COL()
# ASSUME: data rows run from row 2 to the last row that holds a Batch No; the formula is written in the Reorder cell of each of them
# ASSUME: the formula shows YES when Qty On Hand and Reorder Level are both numbers and Qty On Hand is at or below Reorder Level, and "" in every other case
# ASSUME: the Reorder column is not formatted as Text
GOAL     Put a formula in the Reorder column of every data row in the Consumable Stock sheet that shows YES when the quantity on hand is at or below the reorder level.
PROC     AddReorderFormulas
KIND     Function
MODULE   modStock
ARGS     ws As Worksheet
RET      Long
VAR      n As Long
PRE      COL($ws, "Batch No") = 1; COL($ws, "Qty On Hand") = 2; COL($ws, "Reorder Level") = 3; COL($ws, "Reorder") = 4
MODIFIES $ws["Reorder"]
OP       LET(n, LASTROW($ws, 1))
OP       EXIT_IF n < 2
OP       SET_FORMULA(RANGE($ws, "D2:D" & n), "=IF(AND(ISNUMBER(RC2),ISNUMBER(RC3)),IF(RC2<=RC3,""YES"",""""),"""")", style=r1c1)
OP       RETURN(n - 1)
ERR      standard
FIXTURE  stock_01.xlsx sheet=Stock cols=A:Batch No,B:Qty On Hand,C:Reorder Level,D:Reorder rows=6 data=A:seq(1,1);B:int(0,50);C:int(5,20);D:values("old"x6)
FIXTURE  stock_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=stock_01.xlsx call=AddReorderFormulas(Stock) runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND FORMULA(Stock![Reorder]2) = "=IF(AND(ISNUMBER(RC2),ISNUMBER(RC3)),IF(RC2<=RC3,""YES"",""""),"""")" AND FORMULA(Stock![Reorder]7) = "=IF(AND(ISNUMBER(RC2),ISNUMBER(RC3)),IF(RC2<=RC3,""YES"",""""),"""")" AND FORMULA(Stock![Reorder]1) = "" AND FORMULA(Stock![Reorder]8) = ""
TEST     fixture=stock_01.xlsx call=AddReorderFormulas(Stock) assert=ERROR_COUNT = 0 AND COUNTA(Stock![Batch No]) = PRE_COUNTA(Stock![Batch No]) AND SUM(Stock![Qty On Hand]) = PRE_SUM(Stock![Qty On Hand]) AND SUM(Stock![Reorder Level]) = PRE_SUM(Stock![Reorder Level]) AND ROWS_USED(Stock) = PRE_ROWS_USED(Stock)
TEST     fixture=stock_01.xlsx call=AddReorderFormulas(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
