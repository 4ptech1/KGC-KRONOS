VCL 0.6
TASK
# ASSUME: BOM headers in row 1 include "Part No", "Description" and "Unit"
# ASSUME: units are matched ignoring case and stray spaces at the ends; kg and kgs become KG, nos, no and pcs become NOS, mtr, m and meter become MTR
# ASSUME: any other unit and every blank cell are left untouched; a cell already holding exactly the standard text is not counted
GOAL     Standardise the Unit column of the BOM to KG, NOS and MTR and return how many cells were changed.
PROC     StandardiseBomUnits
KIND     Function
MODULE   modBom
ARGS     ws As Worksheet
RET      Long
VAR      cUnit As Long
VAR      changed As Long
VAR      std As String
PRE      COL($ws, "Part No") > 0; COL($ws, "Description") > 0; COL($ws, "Unit") > 0
MODIFIES $ws["Unit"]
OP       LET(cUnit, COL($ws, "Unit"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cUnit) var=r
OP         SKIP_IF IS_ERROR(CELL($ws, r, cUnit))
OP         LET(std, "")
OP         SELECT_CASE UPPER(TRIM(CELL($ws, r, cUnit)))
OP           CASE "KG", "KGS"
OP             LET(std, "KG")
OP           CASE "NOS", "NO", "PCS"
OP             LET(std, "NOS")
OP           CASE "MTR", "M", "METER"
OP             LET(std, "MTR")
OP         IF std <> ""
OP           IF std <> CELL($ws, r, cUnit)
OP             SET_VALUE(CELL($ws, r, cUnit), std)
OP             LET(changed, changed + 1)
OP       RETURN(changed)
PERF     screen_off
ERR      standard
FIXTURE  bom_01.xlsx sheet=Bom cols=A:Part No,B:Description,C:Unit rows=16 data=A:seq(1,1);B:text(10);C:values("kg"x1,"Kg"x1,"kgs"x1,"KG"x1,"nos"x1,"no"x1,"pcs"x1,"NOS"x1,"mtr"x1,"m"x1,"meter"x1,"MTR"x1," Kg "x1,"LTR"x1,""x1,"  M"x1)+boundary("kg","Kg","kgs","KG","nos","no","pcs","NOS","mtr","m","meter","MTR"," Kg ","LTR","","  M")
FIXTURE  bom_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=bom_01.xlsx call=StandardiseBomUnits(Bom) runs=2 assert=ERROR_COUNT = 0 AND COUNTIF(Bom![Unit], "KG") = 5 AND COUNTIF(Bom![Unit], "NOS") = 4 AND COUNTIF(Bom![Unit], "MTR") = 5 AND VALUE(Bom![Unit]2) = "KG" AND VALUE(Bom![Unit]3) = "KG" AND VALUE(Bom![Unit]14) = "KG" AND VALUE(Bom![Unit]15) = "LTR" AND VALUE(Bom![Unit]17) = "MTR" AND VALUE(Bom![Unit]10) = "MTR"
TEST     fixture=bom_01.xlsx call=StandardiseBomUnits(Bom) assert=RESULT = 11 AND ERROR_COUNT = 0 AND COUNTA(Bom![Unit]) = PRE_COUNTA(Bom![Unit]) AND COUNTA(Bom![Part No]) = PRE_COUNTA(Bom![Part No]) AND COUNTA(Bom![Description]) = PRE_COUNTA(Bom![Description])
TEST     fixture=bom_01.xlsx call=StandardiseBomUnits(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
