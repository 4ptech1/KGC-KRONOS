VCL 0.6
TASK
# ASSUME: the BOM sheet is passed in as an argument and is part of the workbook that is passed in
# ASSUME: the copy goes after the last sheet and holds everything the BOM sheet holds
# ASSUME: if a sheet called BOM Copy is already there (any case) it is left exactly as it is, and the macro returns 0
GOAL     Add a sheet called BOM Copy holding a copy of the BOM sheet, only when the workbook has no sheet of that name yet, and return 1 when added.
PROC     AddBomCopySheet
KIND     Function
MODULE   modBom
ARGS     book As Workbook, bom As Worksheet
RET      Long
VAR      added As Long
MODIFIES $book sheets
OP       IF NOT SHEET_EXISTS($book, "BOM Copy")
OP         COPY_SHEET($bom, name="BOM Copy")
OP         LET(added, 1)
OP       RETURN(added)
ERR      standard
FIXTURE  bom_01.xlsx sheet=BOM cols=A:Part No,B:Description,C:Required Qty,D:Unit rows=6 data=A:seq(1,1);B:text(10);C:int(1,50);D:values("EA"x4,"M"x2)
TEST     fixture=bom_01.xlsx call=AddBomCopySheet(bom_01.xlsx, BOM) assert=RESULT = 1 AND ERROR_COUNT = 0 AND SHEET_EXISTS("BOM Copy") AND ROWS_USED('BOM Copy') = ROWS_USED(BOM) AND COUNTA('BOM Copy'![Part No]) = COUNTA(BOM![Part No]) AND SUM('BOM Copy'![Required Qty]) = SUM(BOM![Required Qty])
TEST     fixture=bom_01.xlsx call=AddBomCopySheet(bom_01.xlsx, BOM) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND SHEET_EXISTS("BOM Copy") AND COUNTA(BOM![Part No]) = PRE_COUNTA(BOM![Part No]) AND SUM(BOM![Required Qty]) = PRE_SUM(BOM![Required Qty])
