VCL 0.6
TASK
# ASSUME: the NDE Report sheet has a "Result" header in row 1, and its last row is found from the Result column
# ASSUME: only the Result column is touched; a cell is rewritten only when trimming and upper casing changes it
# ASSUME: a Result cell that holds only spaces is trimmed to empty
GOAL     Trim the spaces from every Result cell in the NDE Report and make it upper case, changing nothing else, and return how many cells changed.
PROC     CleanNdeResults
KIND     Function
MODULE   modNdeReport
ARGS     ws As Worksheet
RET      Long
VAR      cResult As Long
VAR      changed As Long
PRE      COL($ws, "Result") > 0
MODIFIES $ws["Result"]
OP       LET(cResult, COL($ws, "Result"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cResult) var=r
OP         IF UPPER(TRIM(CELL($ws, r, cResult))) <> CELL($ws, r, cResult)
OP           SET_VALUE(CELL($ws, r, cResult), UPPER(TRIM(CELL($ws, r, cResult))))
OP           LET(changed, changed + 1)
OP       RETURN(changed)
ERR      standard
FIXTURE  nde_01.xlsx sheet=NdeReport cols=A:Report No,B:Method,C:Result,D:Inspector rows=15 data=A:seq(1,1);B:values(" rt "x4,"UT"x5,"mt"x6);C:values(" pass"x2,"Fail "x1,"  FAIL  "x1,"pAsS"x1,"PASS"x5,"FAIL"x2,""x2,"   "x1)+boundary(" pass","Fail ","  FAIL  ","pAsS","PASS","FAIL"," pass","","   ");D:text(6)
FIXTURE  nde_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=nde_01.xlsx call=CleanNdeResults(NdeReport) runs=2 assert=VALUE(NdeReport![Result]2) = "PASS" AND VALUE(NdeReport![Result]3) = "FAIL" AND VALUE(NdeReport![Result]4) = "FAIL" AND VALUE(NdeReport![Result]5) = "PASS" AND VALUE(NdeReport![Result]8) = "PASS" AND COUNTIF(NdeReport![Result], "PASS") = 8 AND COUNTIF(NdeReport![Result], "FAIL") = 4 AND ERROR_COUNT = 0
TEST     fixture=nde_01.xlsx call=CleanNdeResults(NdeReport) assert=RESULT = 6 AND VALUE(NdeReport![Result]10) = "" AND VALUE(NdeReport![Result]6) = "PASS" AND VALUE(NdeReport![Result]7) = "FAIL" AND COUNTA(NdeReport![Result]) = PRE_COUNTA(NdeReport![Result]) - 1 AND COUNTA(NdeReport![Report No]) = PRE_COUNTA(NdeReport![Report No]) AND COUNTIF(NdeReport![Method], " rt ") = PRE_COUNTIF(NdeReport![Method], " rt ")
TEST     fixture=nde_01.xlsx call=CleanNdeResults(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
