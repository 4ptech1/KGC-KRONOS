VCL 0.6
TASK
# ASSUME: GRN headers: A "GRN No", B "Part No", C "Ordered Qty", D "Received Qty", E "Receipt Status"; the formula is written as R1C1 so it checks the fixed columns in PRE
# ASSUME: data rows run from row 2 to the last row of GRN No; the formula shows Complete when Received Qty >= Ordered Qty, else Open
GOAL     Write a formula in the Receipt Status column of every GRN row that shows Complete when received quantity reaches ordered quantity and Open otherwise.
PROC     SetReceiptStatusFormulas
KIND     Function
MODULE   modGrn
ARGS     ws As Worksheet
RET      Long
VAR      cNo As Long
VAR      cStatus As Long
VAR      last As Long
VAR      n As Long
PRE      COL($ws, "Ordered Qty") = 3; COL($ws, "Received Qty") = 4; COL($ws, "Receipt Status") = 5; COL($ws, "GRN No") = 1
MODIFIES $ws["Receipt Status"]
OP       LET(cNo, COL($ws, "GRN No"))
OP       LET(cStatus, COL($ws, "Receipt Status"))
OP       LET(last, LASTROW($ws, cNo))
OP       LOOP_ROWS dir=down from=2 to=last var=r
OP         SET_FORMULA(CELL($ws, r, cStatus), "=IF(RC4>=RC3,""Complete"",""Open"")", style=r1c1)
OP         LET(n, n + 1)
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  grn_01.xlsx sheet=GRN cols=A:GRN No,B:Part No,C:Ordered Qty,D:Received Qty,E:Receipt Status rows=6 data=A:seq(1,1);B:text(6);C:int(1,50);D:int(1,50);E:values("old"x6)
FIXTURE  grn_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=grn_01.xlsx call=SetReceiptStatusFormulas(GRN) runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND FORMULA(GRN![Receipt Status]2) = "=IF(RC4>=RC3,""Complete"",""Open"")" AND FORMULA(GRN![Receipt Status]7) = "=IF(RC4>=RC3,""Complete"",""Open"")"
TEST     fixture=grn_01.xlsx call=SetReceiptStatusFormulas(GRN) assert=COUNTA(GRN![GRN No]) = PRE_COUNTA(GRN![GRN No]) AND COUNTA(GRN![Received Qty]) = PRE_COUNTA(GRN![Received Qty]) AND FORMULA(GRN![Receipt Status]1) = "" AND ERROR_COUNT = 0
TEST     fixture=grn_01.xlsx call=SetReceiptStatusFormulas(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
