VCL 0.6
TASK
# ASSUME: BOM headers in row 1 include "Part No" and "Required Qty"; the Flag heading is added after the last heading when it is missing
# ASSUME: a row is flagged when Required Qty is blank, not a number, zero or negative; every other row gets a blank Flag
# ASSUME: rows run from row 2 to the last row that holds a Part No
GOAL     Write Check in the Flag column of every BOM row whose required quantity is blank, not a number, zero or negative, clear the other rows, and return how many were flagged.
PROC     FlagBadBomQty
KIND     Function
MODULE   modBomCheck
ARGS     ws As Worksheet
RET      Long
VAR      cPart As Long
VAR      cQty As Long
VAR      cFlag As Long
VAR      flagged As Long
PRE      COL($ws, "Part No") > 0; COL($ws, "Required Qty") > 0
MODIFIES $ws
OP       LET(cPart, COL($ws, "Part No"))
OP       LET(cQty, COL($ws, "Required Qty"))
OP       LET(cFlag, COL($ws, "Flag"))
OP       IF cFlag = 0
OP         LET(cFlag, LASTCOL($ws, 1) + 1)
OP         SET_VALUE(CELL($ws, 1, cFlag), "Flag")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cPart) var=r
OP         IF IS_NUMBER(CELL($ws, r, cQty))
OP           IF CELL($ws, r, cQty) <= 0
OP             SET_VALUE(CELL($ws, r, cFlag), "Check")
OP             LET(flagged, flagged + 1)
OP           ELSE
OP             SET_VALUE(CELL($ws, r, cFlag), "")
OP         ELSE
OP           SET_VALUE(CELL($ws, r, cFlag), "Check")
OP           LET(flagged, flagged + 1)
OP       RETURN(flagged)
PERF     screen_off
ERR      standard
FIXTURE  bom_01.xlsx sheet=BOM cols=A:Part No,B:Description,C:Required Qty,D:Unit,E:Flag rows=10 data=A:seq(1,1);B:text(10);C:num(1,50)+boundary("",0,-3,0.5,1,2);D:values("EA"x6,"M"x4);E:values("Check"x10)
FIXTURE  bom_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  bom_02.xlsx sheet=BOM cols=A:Part No,B:Description,C:Required Qty,D:Unit rows=10 data=A:seq(1,1);B:text(10);C:values("n/a"x1,"  "x1,""x1,"0"x1,"-3"x1,"1"x1,"2.5"x1,"12"x3)+boundary("n/a","  ","","0","-3","1","2.5","12","12","12");D:values(""x1,"EA"x5,"M"x4)+boundary("","EA","M","EA","M","EA","M","EA","EA","M")
TEST     fixture=bom_01.xlsx call=FlagBadBomQty(BOM) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(BOM![Flag], "Check") = 3 AND COUNTA(BOM![Flag]) = 4 AND VALUE(BOM![Flag]2) = "Check" AND VALUE(BOM![Flag]4) = "Check"
TEST     fixture=bom_02.xlsx call=FlagBadBomQty(BOM) assert=RESULT = 5 AND ERROR_COUNT = 0 AND VALUE(BOM![Flag]1) = "Flag" AND VALUE(BOM!E1) = "Flag" AND VALUE(BOM!E2) = "Check" AND VALUE(BOM!D1) = "Unit" AND COUNTIF(BOM![Flag], "Check") = 5 AND COUNTA(BOM![Flag]) = 6 AND COUNTA(BOM![Part No]) = PRE_COUNTA(BOM![Part No]) AND COUNTA(BOM![Required Qty]) = PRE_COUNTA(BOM![Required Qty]) AND COUNTA(BOM![Unit]) = PRE_COUNTA(BOM![Unit])
TEST     fixture=bom_01.xlsx call=FlagBadBomQty(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(BOM![Flag]) = PRE_COUNTA(BOM![Flag])
