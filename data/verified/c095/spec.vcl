VCL 0.6
TASK
# ASSUME: NCR Register headers: A "NCR No", B "Issue Date", C "Close Date", D "Days Open"; the formulas use these fixed columns, checked in PRE
# ASSUME: a row with no Issue Date gets an empty result; the Days Open cells are already formatted as numbers, not Text
# ASSUME: data rows run from row 2 to the last row of NCR No
GOAL     Write a formula in Days Open for every NCR row: close date minus issue date when closed, otherwise today minus issue date.
PROC     SetDaysOpenFormulas
KIND     Function
MODULE   modNcr
ARGS     ws As Worksheet
RET      Long
VAR      cNo As Long
VAR      cDays As Long
VAR      n As Long
PRE      COL($ws, "NCR No") = 1; COL($ws, "Issue Date") = 2; COL($ws, "Close Date") = 3; COL($ws, "Days Open") = 4
MODIFIES $ws["Days Open"]
OP       LET(cNo, COL($ws, "NCR No"))
OP       LET(cDays, COL($ws, "Days Open"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cNo) var=r
OP         SET_FORMULA(CELL($ws, r, cDays), "=IF(RC2="""","""",IF(RC3<>"""",RC3-RC2,TODAY()-RC2))", style=r1c1)
OP         LET(n, n + 1)
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  ncr_01.xlsx sheet=Ncr cols=A:NCR No,B:Issue Date,C:Close Date,D:Days Open rows=6 data=A:seq(1,1);B:reldates(-60,-1);C:reldates(-30,0)+boundary("",5);D:values("old"x6)
FIXTURE  ncr_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=ncr_01.xlsx call=SetDaysOpenFormulas(Ncr) runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND FORMULA(Ncr![Days Open]2) = "=IF(RC2="""","""",IF(RC3<>"""",RC3-RC2,TODAY()-RC2))" AND FORMULA(Ncr![Days Open]7) = "=IF(RC2="""","""",IF(RC3<>"""",RC3-RC2,TODAY()-RC2))"
TEST     fixture=ncr_01.xlsx call=SetDaysOpenFormulas(Ncr) assert=COUNTA(Ncr![NCR No]) = PRE_COUNTA(Ncr![NCR No]) AND COUNTA(Ncr![Issue Date]) = PRE_COUNTA(Ncr![Issue Date]) AND COUNTA(Ncr![Close Date]) = PRE_COUNTA(Ncr![Close Date]) AND FORMULA(Ncr![Days Open]1) = "" AND ERROR_COUNT = 0
TEST     fixture=ncr_01.xlsx call=SetDaysOpenFormulas(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
