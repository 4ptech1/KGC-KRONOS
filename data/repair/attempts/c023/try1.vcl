VCL 0.6
TASK
# ASSUME: BOM headers (row 1) include "Part No"
# ASSUME: Part No is matched after trimming both ends and ignoring case; the first row of each part stays, with its Part No as typed
# ASSUME: DEDUPE is not used because it does not trim its keys; duplicates are found with a dictionary of trimmed Part Nos instead
# ASSUME: a row with a blank Part No is never a duplicate and is never removed
GOAL     Remove duplicate rows from the BOM sheet by Part No, keeping the first, and return how many duplicate rows were removed.
PROC     RemoveDuplicateParts
KIND     Function
MODULE   modCleanup
ARGS     ws As Worksheet
RET      Long
VAR      cPart As Long
VAR      seen As Dictionary
VAR      copies As Long
VAR      removed As Long
PRE      COL($ws, "Part No") > 0
MODIFIES $ws
OP       LET(cPart, COL($ws, "Part No"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cPart) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cPart))
OP         DICT_ADD(seen, TRIM(CELL($ws, r, cPart)), 1, mode=sum)
OP       LOOP_ROWS dir=up from=LASTROW($ws, cPart) to=2 var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cPart))
OP         DICT_GET(seen, TRIM(CELL($ws, r, cPart)), into=copies)
OP         IF IS_NUMBER(copies)
OP           IF copies > 1
OP             DICT_ADD(seen, TRIM(CELL($ws, r, cPart)), -1, mode=sum)
OP             DELETE_ROW($ws, r)
OP             LET(removed, removed + 1)
OP       RETURN(removed)
PERF     screen_off
ERR      standard
FIXTURE  bom_01.xlsx sheet=BOM cols=A:Part No,B:Description,C:Required Qty rows=10 data=A:values("P-100"x2," P-100 "x1,"P-200"x1,""x2,"p-200"x1,"P-300"x1," P-300"x1,"P-400"x1)+boundary("P-100"," P-100 ","P-200","","P-100","p-200","P-300",""," P-300","P-400");B:values("D-1"x1,"D-2"x1,"D-3"x1,"D-4"x1,"D-5"x1,"D-6"x1,"D-7"x1,"D-8"x1,"D-9"x1,"D-10"x1)+boundary("D-1","D-2","D-3","D-4","D-5","D-6","D-7","D-8","D-9","D-10");C:int(1,50)
FIXTURE  bom_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=bom_01.xlsx call=RemoveDuplicateParts(BOM) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND ROWS_USED(BOM) = 7 AND COUNTA(BOM![Part No]) = 5 AND COUNTIF(BOM![Description], "D-9") = 0
TEST     fixture=bom_01.xlsx call=RemoveDuplicateParts(BOM) assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(BOM![Part No]) = PRE_COUNTA(BOM![Part No]) - RESULT AND COUNTA(BOM![Description]) = PRE_COUNTA(BOM![Description]) - RESULT AND VALUE(BOM![Description]2) = "D-1" AND VALUE(BOM![Description]3) = "D-3" AND VALUE(BOM![Description]4) = "D-4" AND VALUE(BOM![Description]5) = "D-7" AND VALUE(BOM![Description]6) = "D-8" AND VALUE(BOM![Description]7) = "D-10"
TEST     fixture=bom_01.xlsx call=RemoveDuplicateParts(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
