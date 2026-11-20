VCL 0.7
TASK
# ASSUME: Design pressure is typed in barg; the price is the cost with a 12% margin, rounded up to the next 1,000, and never below 50,000.
GOAL     Fill each item's design pressure in kg/cm2 and its selling price, rounded as the estimate rounds them.
PROC     FillRegisterValues
KIND     Function
MODULE   modRegister
ARGS     ws As Worksheet
RET      Long
VAR      n As Long
VAR      cBar As Long
VAR      cKg As Long
VAR      cCost As Long
VAR      cPrice As Long
PRE      COL($ws, "Barg") > 0; COL($ws, "KgCm2") > 0; COL($ws, "Cost") > 0; COL($ws, "Price") > 0
MODIFIES $ws["KgCm2", "Price"]
OP       LET(cBar, COL($ws, "Barg"))
OP       LET(cKg, COL($ws, "KgCm2"))
OP       LET(cCost, COL($ws, "Cost"))
OP       LET(cPrice, COL($ws, "Price"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cCost) var=r
OP         IF IS_NUMBER(CELL($ws, r, cBar))
OP           SET_VALUE(CELL($ws, r, cKg), ROUND(CELL($ws, r, cBar) * 1.01972, 2))
OP           LET(n, n + 1)
OP         IF IS_NUMBER(CELL($ws, r, cCost))
OP           SET_VALUE(CELL($ws, r, cPrice), MAX(ROUNDUP(CELL($ws, r, cCost) * 1.12, -3), 50000))
OP       RETURN(n)
ERR      standard
FIXTURE  register_01.xlsx sheet=Register cols=A:Item,B:Barg,C:KgCm2,D:Cost,E:Price rows=8 data=A:seq(1,1);B:values("18"x1,"35"x2,"69"x2,"77.1"x1,"n/a"x1,""x1)+boundary("18","77.1");D:int(30000,400000)+boundary(40000,288364,289000)
FIXTURE  register_01.xlsx sheet=Notes cols=A:Note rows=1 data=A:values("x"x1)
TEST     fixture=register_01.xlsx call=FillRegisterValues(Register) runs=2 assert=RESULT = 6 AND VALUE(Register!C2) = 18.35 AND VALUE(Register!C3) = 78.62 AND VALUE(Register!E2) = 50000 AND VALUE(Register!E3) = 323000 AND VALUE(Register!E4) = 324000 AND ERROR_COUNT = 0
TEST     fixture=register_01.xlsx call=FillRegisterValues(Register) assert=COUNTA(Register![Barg]) = PRE_COUNTA(Register![Barg]) AND COUNTIF(Register![Price], ">=50000") = 8
TEST     fixture=register_01.xlsx call=FillRegisterValues(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
