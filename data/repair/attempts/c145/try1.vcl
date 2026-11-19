VCL 0.6
TASK
# ASSUME: BOM headers in row 1 include "Part No", "Required Qty", "Issued Qty" and "Short Qty"
# ASSUME: Short Qty is the fourth column, Required Qty the second and Issued Qty the third (the formula text is fixed)
# ASSUME: every row from row 2 down to the last Part No gets the formula; formulas stay live in Excel
GOAL     Put a formula in Short Qty on every BOM data row that gives Required Qty minus Issued Qty when positive, otherwise 0.
PROC     FillShortQtyFormula
KIND     Function
MODULE   modBom
ARGS     ws As Worksheet
RET      Long
VAR      cPart As Long
VAR      last As Long
PRE      COL($ws, "Part No") = 1; COL($ws, "Required Qty") = 2; COL($ws, "Issued Qty") = 3; COL($ws, "Short Qty") = 4
MODIFIES $ws["Short Qty"]
OP       LET(cPart, COL($ws, "Part No"))
OP       LET(last, LASTROW($ws, cPart))
OP       IF last < 2 THEN RETURN(0)
OP       SET_FORMULA(RANGE($ws, "D2:D" & last), "=IF(RC[-2]-RC[-1]>0,RC[-2]-RC[-1],0)", style=r1c1)
OP       RETURN(last - 1)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  bom_01.xlsx sheet=Bom cols=A:Part No,B:Required Qty,C:Issued Qty,D:Short Qty rows=6 data=A:seq(1,1);B:int(1,20);C:int(0,20);D:text(4)
FIXTURE  bom_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=bom_01.xlsx call=FillShortQtyFormula(Bom) runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND FORMULA(Bom![Short Qty]2) = "=IF(RC[-2]-RC[-1]>0,RC[-2]-RC[-1],0)" AND FORMULA(Bom![Short Qty]7) = "=IF(RC[-2]-RC[-1]>0,RC[-2]-RC[-1],0)"
TEST     fixture=bom_01.xlsx call=FillShortQtyFormula(Bom) assert=COUNTA(Bom![Part No]) = PRE_COUNTA(Bom![Part No]) AND COUNTA(Bom![Required Qty]) = PRE_COUNTA(Bom![Required Qty]) AND FORMULA(Bom![Short Qty]1) = ""
TEST     fixture=bom_01.xlsx call=FillShortQtyFormula(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
