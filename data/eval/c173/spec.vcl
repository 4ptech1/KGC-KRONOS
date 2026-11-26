VCL 0.7
TASK
# ASSUME: the register's headings are in row 1 and include "Item" and "PWHT"
# ASSUME: an item row is a row whose Item cell is not blank; rows without an item are left alone
# ASSUME: only a blank PWHT is filled; any text already there (Yes, No, ...) is kept
# ASSUME: the practice fixture is a plain sheet with the register's Item and PWHT headings, because the template has no blank PWHT
GOAL     Write No in the PWHT column of every item row of the enquiry register where PWHT is blank and return how many were filled.
PROC     FillBlankPwht
KIND     Function
MODULE   modRegister
ARGS     ws As Worksheet
RET      Long
VAR      cItem As Long
VAR      cPwht As Long
VAR      filled As Long
PRE      COL($ws, "Item") > 0; COL($ws, "PWHT") > 0
MODIFIES $ws["PWHT"]
OP       LET(cItem, COL($ws, "Item"))
OP       LET(cPwht, COL($ws, "PWHT"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cItem) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cItem))
OP         IF IS_EMPTY(CELL($ws, r, cPwht))
OP           SET_VALUE(CELL($ws, r, cPwht), "No")
OP           LET(filled, filled + 1)
OP       RETURN(filled)
PERF     screen_off
ERR      standard
FIXTURE  register_01.xlsx sheet=Sheet1 cols=A:Item,B:PWHT rows=8 data=A:values("Bottle"x5,"Drum"x2,""x1)+boundary("Bottle","","Drum","Bottle");B:values("Yes"x2,"No"x2,""x3,"Ready"x1)+boundary("","","","Yes","No")
FIXTURE  register_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=register_01.xlsx call=FillBlankPwht(Sheet1) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(Sheet1![PWHT], "No") = 4
TEST     fixture=register_01.xlsx call=FillBlankPwht(Sheet1) assert=RESULT = 2 AND ERROR_COUNT = 0 AND COUNTA(Sheet1![Item]) = PRE_COUNTA(Sheet1![Item]) AND COUNTIF(Sheet1![PWHT], "Yes") = PRE_COUNTIF(Sheet1![PWHT], "Yes")
TEST     fixture=register_01.xlsx call=FillBlankPwht(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
