VCL 0.6
TASK
# ASSUME: BOM headers in row 1 include "Material Grade" and "Required Qty"
# ASSUME: Grade Totals is an existing sheet that is rebuilt on every run
# ASSUME: grades are matched ignoring case and stray spaces, and shown as first typed, trimmed
# ASSUME: a row with no grade is ignored and a Required Qty that is not a number adds nothing
GOAL     Rebuild the Grade Totals sheet with each material grade and its total required quantity from the BOM, biggest total first.
PROC     BuildGradeTotals
KIND     Function
MODULE   modBom
ARGS     bom As Worksheet, out As Worksheet
RET      Long
VAR      cGrade As Long
VAR      cQty As Long
VAR      totals As Dictionary
VAR      total As Double
VAR      outRow As Long
PRE      COL($bom, "Material Grade") > 0; COL($bom, "Required Qty") > 0
MODIFIES $out
OP       LET(cGrade, COL($bom, "Material Grade"))
OP       LET(cQty, COL($bom, "Required Qty"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Material Grade", "Total Required Qty")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($bom, cGrade) var=r
OP         SKIP_IF IS_EMPTY(CELL($bom, r, cGrade))
OP         IF IS_NUMBER(CELL($bom, r, cQty))
OP           DICT_ADD(totals, TRIM(CELL($bom, r, cGrade)), CELL($bom, r, cQty), mode=sum)
OP         ELSE
OP           DICT_ADD(totals, TRIM(CELL($bom, r, cGrade)), 0, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(totals)
OP         DICT_GET(totals, k, into=total)
OP         WRITE_ROW($out, outRow, k, total)
OP         LET(outRow, outRow + 1)
OP       IF outRow > 2
OP         SORT(USED_RANGE($out), key=2, order=desc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  bom_01.xlsx sheet=Bom cols=A:Part No,B:Material Grade,C:Required Qty rows=10 data=A:seq(1,1);B:text(6)+boundary("SA-516 70"," sa-516 70 ","SA-106 B","","sa-106 b","SA-240 304","A105 ","a105","SA-240 304","   ");C:text(3)+boundary(5,7,20,9,"abc","3",4,1,6,8)
FIXTURE  bom_01.xlsx sheet=Totals cols=A:Material Grade,B:Total Required Qty rows=0 stale=3
TEST     fixture=bom_01.xlsx call=BuildGradeTotals(Bom, Totals) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(Totals![Material Grade]) = 5 AND COUNTIF(Totals![Material Grade], "stale") = 0 AND VALUE(Totals![Total Required Qty]2) = 20 AND VALUE(Totals![Total Required Qty]3) = 12 AND VALUE(Totals![Total Required Qty]4) = 9 AND VALUE(Totals![Total Required Qty]5) = 5
TEST     fixture=bom_01.xlsx call=BuildGradeTotals(Bom, Totals) assert=ERROR_COUNT = 0 AND VALUE(Totals![Material Grade]2) = "SA-106 B" AND VALUE(Totals![Material Grade]3) = "SA-516 70" AND VALUE(Totals![Material Grade]4) = "SA-240 304" AND VALUE(Totals![Material Grade]5) = "A105" AND COUNTA(Bom![Part No]) = PRE_COUNTA(Bom![Part No]) AND SUM(Bom![Required Qty]) = PRE_SUM(Bom![Required Qty]) AND COUNTA(Bom![Material Grade]) = PRE_COUNTA(Bom![Material Grade])
TEST     fixture=bom_01.xlsx call=BuildGradeTotals(Totals, Bom) assert=RESULT = -1 AND ERROR_COUNT = 1
