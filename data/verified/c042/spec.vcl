VCL 0.6
TASK
# ASSUME: Grn headers in row 1 include "GRN No" and "Part No"; no other column is touched
# ASSUME: trimming removes only the spaces at both ends of a cell; spaces inside the text stay
# ASSUME: Part No is trimmed and made upper case; GRN No is only trimmed and keeps its case
# ASSUME: a cell counts as changed only when its text differs from what it held; a cell holding a number is not changed
GOAL     Trim the spaces from the Part No and GRN No cells of the GRN sheet, make Part No upper case, and return how many cells changed.
PROC     TidyGrnCells
KIND     Function
MODULE   modGrnCleanup
ARGS     grn As Worksheet
RET      Long
VAR      cPart As Long
VAR      cGrn As Long
VAR      newPart As String
VAR      newGrn As String
VAR      changed As Long
PRE      COL($grn, "Part No") > 0; COL($grn, "GRN No") > 0
MODIFIES $grn["Part No", "GRN No"]
OP       LET(cPart, COL($grn, "Part No"))
OP       LET(cGrn, COL($grn, "GRN No"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($grn, cPart) var=r
OP         LET(newPart, UPPER(TRIM(CELL($grn, r, cPart))))
OP         IF newPart <> CELL($grn, r, cPart)
OP           SET_VALUE(CELL($grn, r, cPart), newPart)
OP           LET(changed, changed + 1)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($grn, cGrn) var=g
OP         LET(newGrn, TRIM(CELL($grn, g, cGrn)))
OP         IF newGrn <> CELL($grn, g, cGrn)
OP           SET_VALUE(CELL($grn, g, cGrn), newGrn)
OP           LET(changed, changed + 1)
OP       RETURN(changed)
ERR      standard
FIXTURE  grn_01.xlsx sheet=Grn cols=A:GRN No,B:Part No,C:Supplier,D:Qty rows=8 data=A:text(8)+boundary(" GRN-001","GRN-002","GRN-003 ","GRN-004","","GRN-006","GRN-007","  GRN-008  ");B:text(8)+boundary("abc-100"," ABC-200 ","ABC-300","Abc-400","ABC-500","ABC-600","  xyz-7","");C:text(6);D:int(1,50)
FIXTURE  grn_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=grn_01.xlsx call=TidyGrnCells(Grn) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE(Grn![GRN No]2) = "GRN-001" AND VALUE(Grn![Part No]2) = "ABC-100" AND VALUE(Grn![GRN No]9) = "GRN-008"
TEST     fixture=grn_01.xlsx call=TidyGrnCells(Grn) assert=RESULT = 7 AND ERROR_COUNT = 0 AND VALUE(Grn![Part No]3) = "ABC-200" AND VALUE(Grn![GRN No]4) = "GRN-003" AND VALUE(Grn![Part No]5) = "ABC-400" AND VALUE(Grn![Part No]8) = "XYZ-7"
TEST     fixture=grn_01.xlsx call=TidyGrnCells(Grn) assert=COUNTA(Grn![GRN No]) = PRE_COUNTA(Grn![GRN No]) AND COUNTA(Grn![Part No]) = PRE_COUNTA(Grn![Part No]) AND SUM(Grn![Qty]) = PRE_SUM(Grn![Qty]) AND VALUE(Grn![GRN No]3) = "GRN-002"
TEST     fixture=grn_01.xlsx call=TidyGrnCells(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
