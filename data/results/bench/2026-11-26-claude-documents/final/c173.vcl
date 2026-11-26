VCL 0.7
TASK
# ASSUME: headers are in row 1 and include "Item" and "PWHT"; an item row is a row whose Item cell is not blank
# ASSUME: a PWHT cell holding only spaces counts as blank; the other register rows and columns are left alone
# ASSUME: the function returns how many PWHT cells it filled
GOAL     Write No in the PWHT column of every register item row where PWHT is blank and return how many cells were filled.
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
ERR      standard
FIXTURE  register_01.xlsx sheet=Register cols=A:Item,B:Qty,C:PWHT rows=8 data=A:values("Bottle"x6,"Drum"x1,""x1)+boundary("Bottle","Bottle","","Bottle","Bottle");B:int(1,5);C:values(""x3,"Yes"x3,"No"x2)+boundary("","Yes","","No","")
FIXTURE  register_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  register_02.xlsx template=enquiry_register
TEST     fixture=register_01.xlsx call=FillBlankPwht(Register) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(Register![PWHT], "No") = PRE_COUNTIF(Register![PWHT], "No") + 2
TEST     fixture=register_01.xlsx call=FillBlankPwht(Register) assert=RESULT = 2 AND ERROR_COUNT = 0 AND COUNTA(Register![Item]) = PRE_COUNTA(Register![Item]) AND COUNTIF(Register![PWHT], "Yes") = PRE_COUNTIF(Register![PWHT], "Yes") AND VALUE(Register![PWHT]2) = "No"
TEST     fixture=register_01.xlsx call=FillBlankPwht(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
TEST     fixture=register_02.xlsx call=FillBlankPwht(Sheet1) assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(Sheet1![PWHT], "No") = PRE_COUNTIF(Sheet1![PWHT], "No")
