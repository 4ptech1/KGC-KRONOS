VCL 0.7
TASK
# ASSUME: the GRN sheet headings are in row 1; Receipt Status is the 5th column and Received Qty and Ordered Qty are the 4th and 3rd
# ASSUME: the formula is written to every data row down to the last GRN No row, as R1C1 text
GOAL     Put a formula in every data row of the Receipt Status column that shows Complete when Received Qty is at least Ordered Qty and Open otherwise.
PROC     SetReceiptStatusFormulas
KIND     Function
MODULE   modGrn
ARGS     ws As Worksheet
RET      Long
VAR      cGrn As Long
VAR      last As Long
PRE      COL($ws, "GRN No") = 1; COL($ws, "Ordered Qty") = 3; COL($ws, "Received Qty") = 4; COL($ws, "Receipt Status") = 5
MODIFIES $ws["Receipt Status"]
OP       LET(cGrn, COL($ws, "GRN No"))
OP       LET(last, LASTROW($ws, cGrn))
OP       EXIT_IF last < 2
OP       SET_FORMULA(BLOCK($ws, 2, 5, last, 5), "=IF(RC4>=RC3,""Complete"",""Open"")", style=r1c1)
OP       RETURN(last - 1)
ERR      standard
FIXTURE  grn_01.xlsx sheet=Grn cols=A:GRN No,B:Part No,C:Ordered Qty,D:Received Qty,E:Receipt Status rows=6 data=A:seq(1,1);B:text(6);C:int(1,10);D:int(1,10);E:values("old"x6)
FIXTURE  grn_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=grn_01.xlsx call=SetReceiptStatusFormulas(Grn) runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND FORMULA(Grn![Receipt Status]2) = "=IF(RC4>=RC3,""Complete"",""Open"")" AND FORMULA(Grn![Receipt Status]7) = "=IF(RC4>=RC3,""Complete"",""Open"")"
TEST     fixture=grn_01.xlsx call=SetReceiptStatusFormulas(Grn) assert=COUNTA(Grn![GRN No]) = PRE_COUNTA(Grn![GRN No]) AND COUNTA(Grn![Part No]) = PRE_COUNTA(Grn![Part No]) AND FORMULA(Grn![Receipt Status]1) = ""
TEST     fixture=grn_01.xlsx call=SetReceiptStatusFormulas(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
