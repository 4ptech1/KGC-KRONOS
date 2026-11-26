VCL 0.7
TASK
# ASSUME: the Inputs sheet has the headers "Ref" and "Value" in row 1; the real sheet is called Inputs, and the TESTs use the short form Inputs with the template's own headings
# ASSUME: the parameters are found by their Ref text, r_plate and margin, ignoring case and spaces; the margin is written as the fraction 0.12
# ASSUME: a Ref that is not on the sheet is not added; nothing else on the sheet is written
GOAL     Set the plate rate r_plate to 95 and the margin to 0.12 in the Value column of the Inputs sheet and change nothing else.
PROC     SetPlateRateAndMargin
KIND     Function
MODULE   modEstimation
ARGS     ws As Worksheet
RET      Long
VAR      cRef As Long
VAR      cVal As Long
VAR      n As Long
PRE      COL($ws, "Ref") > 0; COL($ws, "Value") > 0
MODIFIES $ws["Value"]
OP       LET(cRef, COL($ws, "Ref"))
OP       LET(cVal, COL($ws, "Value"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cRef) var=r
OP         IF SAME(CELL($ws, r, cRef), "r_plate")
OP           SET_VALUE(CELL($ws, r, cVal), 95)
OP           LET(n, n + 1)
OP         IF SAME(CELL($ws, r, cRef), "margin")
OP           SET_VALUE(CELL($ws, r, cVal), 0.12)
OP           LET(n, n + 1)
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  inputs_01.xlsx sheet=Inputs cols=A:Ref,B:Parameter,C:Value,D:Unit rows=6 data=A:values("fx"x1,"r_plate"x1,"r_forg"x1,"margin"x1,"oh"x1,"cont"x1)+boundary("fx","r_plate","r_forg","margin","oh","cont");B:text(10);C:num(1,200)+boundary(85,80,130,0.05,0.1,0.02);D:values("INR/kg"x6)
FIXTURE  inputs_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=inputs_01.xlsx call=SetPlateRateAndMargin(Inputs) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(Inputs![Value]3) = 95 AND VALUE(Inputs![Value]5) = 0.12
TEST     fixture=inputs_01.xlsx call=SetPlateRateAndMargin(Inputs) assert=VALUE(Inputs![Value]2) = 85 AND VALUE(Inputs![Value]4) = 130 AND VALUE(Inputs![Value]6) = 0.1 AND VALUE(Inputs![Value]7) = 0.02 AND COUNTA(Inputs![Value]) = PRE_COUNTA(Inputs![Value]) AND COUNTA(Inputs![Ref]) = PRE_COUNTA(Inputs![Ref])
TEST     fixture=inputs_01.xlsx call=SetPlateRateAndMargin(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
