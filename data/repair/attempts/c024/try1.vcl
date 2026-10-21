VCL 0.6
TASK
# ASSUME: Material Certs has the header "Heat No" in row 1
# ASSUME: trimmed means no spaces, tabs or line breaks at either end; spaces inside a heat number stay
# ASSUME: a cell that is blank or holds only spaces is left as it is and is not counted
# ASSUME: a cell that is already trimmed and in upper case is not written and not counted
GOAL     Trim every Heat No on the Material Certs sheet and make it upper case, change nothing else and return how many cells changed.
PROC     TidyHeatNumbers
KIND     Function
MODULE   modCleanup
ARGS     ws As Worksheet
RET      Long
VAR      cHeat As Long
VAR      changed As Long
PRE      COL($ws, "Heat No") > 0
MODIFIES $ws["Heat No"]
OP       LET(cHeat, COL($ws, "Heat No"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cHeat) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cHeat))
OP         IF CELL($ws, r, cHeat) <> UPPER(TRIM(CELL($ws, r, cHeat)))
OP           SET_VALUE(CELL($ws, r, cHeat), UPPER(TRIM(CELL($ws, r, cHeat))))
OP           LET(changed, changed + 1)
OP       RETURN(changed)
PERF     screen_off
ERR      standard
FIXTURE  certs_01.xlsx sheet=Certs cols=A:Heat No,B:MTC No rows=8 data=A:values("H-100"x1," h-200 "x1,"h-300"x1,""x1,"H-400 "x1,"   "x1," H-500"x1,"H-600"x1)+boundary("H-100"," h-200 ","h-300","","H-400 ","   "," H-500","H-600");B:values(" m-1 "x1,"m-2"x1,"M-3 "x1," M-4"x1,"m-5"x1,"M-6"x1," m-7"x1,"M-8 "x1)+boundary(" m-1 ","m-2","M-3 "," M-4","m-5","M-6"," m-7","M-8 ")
FIXTURE  certs_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=certs_01.xlsx call=TidyHeatNumbers(Certs) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE(Certs![Heat No]3) = "H-200" AND VALUE(Certs![Heat No]4) = "H-300" AND VALUE(Certs![Heat No]6) = "H-400" AND VALUE(Certs![Heat No]8) = "H-500"
TEST     fixture=certs_01.xlsx call=TidyHeatNumbers(Certs) assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(Certs![Heat No]) = PRE_COUNTA(Certs![Heat No]) AND COUNTA(Certs![MTC No]) = PRE_COUNTA(Certs![MTC No]) AND VALUE(Certs![MTC No]2) = " m-1 " AND VALUE(Certs![MTC No]3) = "m-2" AND VALUE(Certs![Heat No]7) = "   "
TEST     fixture=certs_01.xlsx call=TidyHeatNumbers(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
