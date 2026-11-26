VCL 0.7
TASK
# ASSUME: the Conversion headings are in row 4 (Sl in A to Amount in J), so PRE checks them by their cells
# ASSUME: the new activity goes in the row below the last used row of column A, and its Sl is that row's Sl plus one
# ASSUME: the Total Qty formula is =Qty/set * Sets (E*F) and the Amount formula is =Total Qty * Rate (G*I), as in the row above
# ASSUME: a re-run does nothing once the last activity is already PWHT of floating head
# ASSUME: the Total CH, FH and ALL sums below the list are left as they are
GOAL     Add the PWHT of floating head activity after the last row of the Conversion sheet with the same Total Qty and Amount formulas as the row above.
PROC     AddPwhtActivity
KIND     Function
MODULE   modEstimation
ARGS     ws As Worksheet
RET      Long
VAR      last As Long
VAR      newRow As Long
PRE      CELL($ws, 4, 1) = "Sl"; CELL($ws, 4, 2) = "Scope"; CELL($ws, 4, 3) = "Activity"; CELL($ws, 4, 4) = "Basis / Remarks"; CELL($ws, 4, 5) = "Qty / set"; CELL($ws, 4, 6) = "Sets"; CELL($ws, 4, 7) = "Total Qty"; CELL($ws, 4, 8) = "Unit"; CELL($ws, 4, 9) = "Rate (INR)"; CELL($ws, 4, 10) = "Amount (INR)"; IS_NUMBER(CELL($ws, LASTROW($ws, 1), 1))
MODIFIES $ws
OP       LET(last, LASTROW($ws, 1))
OP       EXIT_IF SAME(CELL($ws, last, 3), "PWHT of floating head")
OP       LET(newRow, last + 1)
OP       SET_VALUE(CELL($ws, newRow, 1), CELL($ws, last, 1) + 1)
OP       SET_VALUE(CELL($ws, newRow, 2), "FH")
OP       SET_VALUE(CELL($ws, newRow, 3), "PWHT of floating head")
OP       SET_VALUE(CELL($ws, newRow, 4), "furnace, 1 lot")
OP       SET_VALUE(CELL($ws, newRow, 5), 1)
OP       SET_VALUE(CELL($ws, newRow, 6), 2)
OP       SET_FORMULA(CELL($ws, newRow, 7), "=RC5*RC6", style=r1c1)
OP       SET_VALUE(CELL($ws, newRow, 8), "Lot")
OP       SET_VALUE(CELL($ws, newRow, 9), 40000)
OP       SET_FORMULA(CELL($ws, newRow, 10), "=RC7*RC9", style=r1c1)
OP       RETURN(newRow)
PERF     screen_off
ERR      standard
FIXTURE  est_01.xlsx template=estimation
TEST     fixture=est_01.xlsx call=AddPwhtActivity(Conversion) assert=RESULT = 9 AND ERROR_COUNT = 0 AND VALUE(Conversion!A9) = 5 AND VALUE(Conversion!B9) = "FH" AND VALUE(Conversion!C9) = "PWHT of floating head" AND VALUE(Conversion!D9) = "furnace, 1 lot" AND VALUE(Conversion!E9) = 1 AND VALUE(Conversion!F9) = 2 AND VALUE(Conversion!H9) = "Lot" AND VALUE(Conversion!I9) = 40000 AND FORMULA(Conversion!G9) = "=RC5*RC6" AND FORMULA(Conversion!J9) = "=RC7*RC9" AND ROWS_USED(Conversion) = 13
TEST     fixture=est_01.xlsx call=AddPwhtActivity(Conversion) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(Conversion![Activity], "PWHT of floating head") = 1 AND VALUE(Conversion!A8) = 4 AND VALUE(Conversion!C8) = "Hydrotest & painting" AND VALUE(Conversion!I8) = 15000 AND VALUE(Conversion!A10) = "" AND VALUE(Conversion!I10) = "Total CH"
TEST     fixture=est_01.xlsx call=AddPwhtActivity('Enq Summary') assert=RESULT = -1 AND ERROR_COUNT = 1
