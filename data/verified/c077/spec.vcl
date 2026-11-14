VCL 0.6
TASK
# ASSUME: BOM headers in row 1 include "Part No", "Unit" and "Remark"
# ASSUME: the rows run down to the last Part No; a row with no Part No is not a BOM line, so it is skipped and its Remark is left as it is
# ASSUME: a Unit is good when it is NOS, KG, MTR or SET, ignoring case and spaces at the ends; anything else, an error value included, gets Unit?
# ASSUME: the number of rows that need a unit fixed is the macro's return value
GOAL     Write Unit? in the Remark of every BOM row whose Unit is blank or not NOS, KG, MTR or SET, clear the Remark of the other rows and return how many rows need a unit fixed.
PROC     FlagBadUnits
KIND     Function
MODULE   modBom
ARGS     ws As Worksheet
RET      Long
VAR      cPart As Long
VAR      cUnit As Long
VAR      cRemark As Long
VAR      n As Long
PRE      COL($ws, "Part No") > 0; COL($ws, "Unit") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Remark"]
OP       LET(cPart, COL($ws, "Part No"))
OP       LET(cUnit, COL($ws, "Unit"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cPart) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cPart))
OP         IF SAME(CELL($ws, r, cUnit), "NOS") OR SAME(CELL($ws, r, cUnit), "KG") OR SAME(CELL($ws, r, cUnit), "MTR") OR SAME(CELL($ws, r, cUnit), "SET")
OP           CLEAR(CELL($ws, r, cRemark), what=contents)
OP         ELSE
OP           SET_VALUE(CELL($ws, r, cRemark), "Unit?")
OP           LET(n, n + 1)
OP       RETURN(n)
ERR      standard
FIXTURE  bom_units_01.xlsx sheet=BOM cols=A:Part No,B:Description,C:Required Qty,D:Unit,E:Remark rows=12 data=A:values("P-101"x1,"P-102"x1,"P-103"x1,"P-104"x1,"P-105"x1,""x2,"P-107"x1,"P-108"x1,"P-109"x1,"P-111"x1,"P-112"x1)+boundary("P-101","P-102","P-103","P-104","P-105","","P-107","P-108","P-109","","P-111","P-112");B:text(10);C:int(1,50);D:values("nos "x1,"kg"x1," Mtr "x1,"set"x1,""x2,"PCS"x2,"KGS"x1,"EA"x1,"NOS"x1,"SET"x1)+boundary("nos ","kg"," Mtr ","set","","PCS","PCS","KGS","EA","NOS","SET","");E:values("old"x4,""x4,"Unit?"x2,"keep"x2)+boundary("old","","Unit?","old","","keep","old","","Unit?","keep","old","")
FIXTURE  bom_units_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=bom_units_01.xlsx call=FlagBadUnits(BOM) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTIF(BOM![Remark], "Unit?") = 5 AND COUNTIF(BOM![Remark], "old") = 0 AND COUNTIF(BOM![Remark], "keep") = 2 AND COUNTA(BOM![Remark]) = 8 AND VALUE(BOM![Remark]6) = "Unit?" AND VALUE(BOM![Remark]8) = "Unit?" AND VALUE(BOM![Remark]9) = "Unit?" AND VALUE(BOM![Remark]10) = "Unit?" AND VALUE(BOM![Remark]13) = "Unit?"
TEST     fixture=bom_units_01.xlsx call=FlagBadUnits(BOM) assert=RESULT = 5 AND COUNTA(BOM![Part No]) = PRE_COUNTA(BOM![Part No]) AND COUNTA(BOM![Unit]) = PRE_COUNTA(BOM![Unit]) AND SUM(BOM![Required Qty]) = PRE_SUM(BOM![Required Qty]) AND VALUE(BOM![Remark]7) = "keep" AND VALUE(BOM![Remark]11) = "keep"
TEST     fixture=bom_units_01.xlsx call=FlagBadUnits(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
