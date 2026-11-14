VCL 0.6
TASK
# ASSUME: BOM headers in row 1 include "Part No", "Required Qty" and "Unit"
# ASSUME: the Unit Totals sheet exists; it is cleared and rebuilt on every run with the headers Unit and Total Required Qty
# ASSUME: the rows run down to the last Part No; each of them is either totalled or counted as left out, a wholly blank row too
# ASSUME: units are trimmed and written in upper case, so KG, kg and " Kg " are one unit, KG
# ASSUME: a Required Qty typed as text that is a plain number ("25") counts as a number
# ASSUME: the number of rows left out is the macro's return value
GOAL     Rebuild the Unit Totals sheet with the total Required Qty of each Unit in the BOM, units A to Z, and return how many rows were left out for a blank unit or a quantity that is not a number.
PROC     BuildUnitTotals
KIND     Function
MODULE   modBom
ARGS     bom As Worksheet, out As Worksheet
RET      Long
VAR      cPart As Long
VAR      cQty As Long
VAR      cUnit As Long
VAR      totals As Dictionary
VAR      got As Double
VAR      outRow As Long
VAR      leftOut As Long
PRE      COL($bom, "Part No") > 0; COL($bom, "Required Qty") > 0; COL($bom, "Unit") > 0
MODIFIES $out
OP       LET(cPart, COL($bom, "Part No"))
OP       LET(cQty, COL($bom, "Required Qty"))
OP       LET(cUnit, COL($bom, "Unit"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($bom, cPart) var=r
OP         IF IS_EMPTY(CELL($bom, r, cUnit))
OP           LET(leftOut, leftOut + 1)
OP         ELSE
OP           IF IS_NUMBER(CELL($bom, r, cQty))
OP             DICT_ADD(totals, UPPER(TRIM(CELL($bom, r, cUnit))), CELL($bom, r, cQty), mode=sum)
OP           ELSE
OP             LET(leftOut, leftOut + 1)
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Unit", "Total Required Qty")
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(totals)
OP         DICT_GET(totals, k, into=got)
OP         WRITE_ROW($out, outRow, k, got)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       RETURN(leftOut)
PERF     screen_off
ERR      standard
FIXTURE  bom_totals_01.xlsx sheet=BOM cols=A:Part No,B:Description,C:Required Qty,D:Unit rows=12 data=A:seq(101,1);B:text(10);C:num(1,20)+boundary(10,5,4,"",7,"n/a",2.5,6,3,8,12,"");D:values("kg"x1,"KG"x1,"NOS"x2,""x2,"MTR"x2," Mtr "x1,"SET"x1,"nos "x1,"Kg"x1)+boundary("kg","KG","NOS","NOS","","MTR"," Mtr ","SET","nos ","Kg","MTR","")
FIXTURE  bom_totals_01.xlsx sheet=UnitTotals cols=A:Unit,B:Total Required Qty rows=0 stale=6
TEST     fixture=bom_totals_01.xlsx call=BuildUnitTotals(BOM, UnitTotals) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND ROWS_USED(UnitTotals) = 5 AND VALUE(UnitTotals![Unit]2) = "KG" AND VALUE(UnitTotals![Total Required Qty]2) = 23 AND VALUE(UnitTotals![Unit]3) = "MTR" AND VALUE(UnitTotals![Total Required Qty]3) = 14.5 AND VALUE(UnitTotals![Unit]4) = "NOS" AND VALUE(UnitTotals![Total Required Qty]4) = 7 AND VALUE(UnitTotals![Unit]5) = "SET" AND VALUE(UnitTotals![Total Required Qty]5) = 6 AND COUNTIF(UnitTotals![Unit], "stale") = 0 AND FILL_COUNT(UnitTotals!A1:B8, #FF0000) = 0
TEST     fixture=bom_totals_01.xlsx call=BuildUnitTotals(BOM, UnitTotals) assert=RESULT = 4 AND COUNTA(BOM![Part No]) = PRE_COUNTA(BOM![Part No]) AND COUNTA(BOM![Unit]) = PRE_COUNTA(BOM![Unit]) AND SUM(BOM![Required Qty]) = PRE_SUM(BOM![Required Qty]) AND SUM(UnitTotals![Total Required Qty]) = 50.5 AND COUNTA(UnitTotals![Unit]) = 5
TEST     fixture=bom_totals_01.xlsx call=BuildUnitTotals(UnitTotals, BOM) assert=RESULT = -1 AND ERROR_COUNT = 1
