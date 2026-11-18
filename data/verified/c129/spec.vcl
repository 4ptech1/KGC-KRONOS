VCL 0.6
TASK
# ASSUME: Material Certs headers in row 1 include "Heat No" and "MTC No"; F1 is free for the answer
# ASSUME: the heat number is given as the String argument heat; it is matched ignoring case and stray spaces
# ASSUME: if a heat number appears more than once, the first row (nearest the top) wins
# ASSUME: a blank heat number is not found; the row returned is the sheet row, or 0
GOAL     Find a heat number on the Material Certs sheet, write its MTC No in cell F1 or Not found, and return its row or 0.
PROC     FindHeatCert
KIND     Function
MODULE   modCerts
ARGS     ws As Worksheet, heat As String
RET      Long
VAR      cHeat As Long
VAR      cMtc As Long
VAR      found As Long
PRE      COL($ws, "Heat No") > 0; COL($ws, "MTC No") > 0
MODIFIES $ws
OP       LET(cHeat, COL($ws, "Heat No"))
OP       LET(cMtc, COL($ws, "MTC No"))
OP       LET(found, 0)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cHeat) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cHeat))
OP         IF SAME(CELL($ws, r, cHeat), $heat)
OP           LET(found, r)
OP           SET_VALUE(CELL($ws, 1, 6), CELL($ws, r, cMtc))
OP         EXIT_LOOP_IF found > 0
OP       IF found = 0 THEN SET_VALUE(CELL($ws, 1, 6), "Not found")
OP       RETURN(found)
ERR      standard
FIXTURE  certs_01.xlsx sheet=Certs cols=A:Heat No,B:MTC No,C:Supplier rows=6 data=A:values("H-100"x1,""x1,"H-200"x1," h-100 "x1,"H-300"x1,"H-400"x1)+boundary("H-100","","H-200"," h-100 ","H-300","H-400");B:values("M-1"x1,"M-2"x1,"M-3"x1,"M-4"x1,"M-5"x1,"M-6"x1)+boundary("M-1","M-2","M-3","M-4","M-5","M-6");C:text(6)
FIXTURE  certs_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=certs_01.xlsx call=FindHeatCert(Certs, "H-100") runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(Certs!F1) = "M-1"
TEST     fixture=certs_01.xlsx call=FindHeatCert(Certs, " h-300 ") assert=RESULT = 6 AND VALUE(Certs!F1) = "M-5" AND COUNTA(Certs![Heat No]) = PRE_COUNTA(Certs![Heat No]) AND COUNTA(Certs![MTC No]) = PRE_COUNTA(Certs![MTC No])
TEST     fixture=certs_01.xlsx call=FindHeatCert(Certs, "H-400") assert=RESULT = 7 AND VALUE(Certs!F1) = "M-6"
TEST     fixture=certs_01.xlsx call=FindHeatCert(Certs, "H-200") assert=RESULT = 4 AND VALUE(Certs!F1) = "M-3"
TEST     fixture=certs_01.xlsx call=FindHeatCert(Certs, "H-999") runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE(Certs!F1) = "Not found"
TEST     fixture=certs_01.xlsx call=FindHeatCert(Certs, "") assert=RESULT = 0 AND VALUE(Certs!F1) = "Not found"
TEST     fixture=certs_01.xlsx call=FindHeatCert(Notes, "H-100") assert=RESULT = -1 AND ERROR_COUNT = 1
