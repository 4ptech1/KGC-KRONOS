VCL 0.6
TASK
# ASSUME: Dispatch headers in row 1 include "Dispatch No", "Planned Date", "Actual Date" and "Days Late"
# ASSUME: Planned Date is column B and Actual Date is column C because a formula's text names fixed columns; the macro stops without changing anything if they are elsewhere
# ASSUME: every row from 2 down to the last Dispatch No is a data row; its Days Late cell gets the formula and loses whatever it held
# ASSUME: every dispatch has a Planned Date; the formula is blank only when the Actual Date is blank
GOAL     Put a formula in the Days Late cell of every data row that gives Actual Date minus Planned Date, or stays blank when there is no Actual Date, and return how many rows got it.
PROC     AddDaysLateFormulas
KIND     Function
MODULE   modDispatch
ARGS     ws As Worksheet
RET      Long
VAR      cNo As Long
VAR      cLate As Long
VAR      filled As Long
PRE      COL($ws, "Dispatch No") > 0; COL($ws, "Planned Date") = 2; COL($ws, "Actual Date") = 3; COL($ws, "Days Late") > 0
MODIFIES $ws["Days Late"]
OP       LET(cNo, COL($ws, "Dispatch No"))
OP       LET(cLate, COL($ws, "Days Late"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cNo) var=r
OP         SET_FORMULA(CELL($ws, r, cLate), "=IF(RC3="""","""",RC3-RC2)", style=r1c1)
OP         LET(filled, filled + 1)
OP       RETURN(filled)
ERR      standard
FIXTURE  dispatch_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Planned Date,C:Actual Date,D:Days Late rows=9 data=A:seq(1,1);B:dates(2026-01-01,2026-09-30);C:dates(2026-01-05,2026-10-15)+blank(3);D:int(0,9)
FIXTURE  dispatch_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  dispatch_02.xlsx sheet=Dispatch cols=A:Dispatch No,B:Actual Date,C:Planned Date,D:Days Late rows=4 data=A:seq(1,1);B:dates(2026-01-05,2026-10-15);C:dates(2026-01-01,2026-09-30);D:int(0,9)
TEST     fixture=dispatch_01.xlsx call=AddDaysLateFormulas(Dispatch) runs=2 assert=RESULT = 9 AND ERROR_COUNT = 0 AND FORMULA(Dispatch![Days Late]2) = "=IF(RC3="""","""",RC3-RC2)" AND FORMULA(Dispatch![Days Late]10) = "=IF(RC3="""","""",RC3-RC2)" AND FORMULA(Dispatch![Days Late]11) = ""
TEST     fixture=dispatch_01.xlsx call=AddDaysLateFormulas(Dispatch) assert=FORMULA(Dispatch![Days Late]3) = "=IF(RC3="""","""",RC3-RC2)" AND FORMULA(Dispatch![Days Late]4) = "=IF(RC3="""","""",RC3-RC2)" AND FORMULA(Dispatch![Days Late]5) = "=IF(RC3="""","""",RC3-RC2)" AND FORMULA(Dispatch![Days Late]6) = "=IF(RC3="""","""",RC3-RC2)" AND FORMULA(Dispatch![Days Late]7) = "=IF(RC3="""","""",RC3-RC2)" AND FORMULA(Dispatch![Days Late]8) = "=IF(RC3="""","""",RC3-RC2)" AND FORMULA(Dispatch![Days Late]9) = "=IF(RC3="""","""",RC3-RC2)"
TEST     fixture=dispatch_01.xlsx call=AddDaysLateFormulas(Dispatch) assert=RESULT = 9 AND COUNTA(Dispatch![Dispatch No]) = PRE_COUNTA(Dispatch![Dispatch No]) AND COUNTA(Dispatch![Planned Date]) = PRE_COUNTA(Dispatch![Planned Date]) AND COUNTA(Dispatch![Actual Date]) = PRE_COUNTA(Dispatch![Actual Date])
TEST     fixture=dispatch_01.xlsx call=AddDaysLateFormulas(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
TEST     fixture=dispatch_02.xlsx call=AddDaysLateFormulas(Dispatch) assert=RESULT = -1 AND ERROR_COUNT = 1 AND FORMULA(Dispatch![Days Late]2) = ""
