VCL 0.6
TASK
# ASSUME: the real sheet is "Consumable Issue"; its row-1 headers are "Batch No", "Opening Stock", "Issued" and "Balance" (the TESTs call it Issue)
# ASSUME: every row from row 2 down that has a Batch No gets the formula; a row with a blank Batch No is left as it is
# ASSUME: the old content of Balance is replaced by the formula, and Balance is not formatted as Text
# ASSUME: Opening Stock and Issued are numbers typed by hand; a bad value shows as an Excel error in Balance and is not hidden
GOAL     Put the formula Opening Stock minus Issued into the Balance cell of every data row of the Consumable Issue sheet.
PROC     FillBalanceFormulas
KIND     Function
MODULE   modConsumables
ARGS     ws As Worksheet
RET      Long
VAR      cBatch As Long
VAR      cBalance As Long
VAR      filled As Long
PRE      COL($ws, "Batch No") > 0; COL($ws, "Opening Stock") = 2; COL($ws, "Issued") = 3; COL($ws, "Balance") > 0
MODIFIES $ws["Balance"]
OP       LET(cBatch, COL($ws, "Batch No"))
OP       LET(cBalance, COL($ws, "Balance"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cBatch) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cBatch))
OP         SET_FORMULA(CELL($ws, r, cBalance), "=RC2-RC3", style=r1c1)
OP         LET(filled, filled + 1)
OP       RETURN(filled)
PERF     screen_off
ERR      standard
FIXTURE  issue_01.xlsx sheet=Issue cols=A:Batch No,B:Opening Stock,C:Issued,D:Balance rows=8 data=A:text(6)+boundary("B-100","","B-102");B:int(50,500);C:int(0,50);D:int(0,99)
FIXTURE  issue_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=issue_01.xlsx call=FillBalanceFormulas(Issue) runs=2 assert=RESULT = 7 AND ERROR_COUNT = 0 AND FORMULA(Issue![Balance]1) = "" AND FORMULA(Issue![Balance]2) = "=RC2-RC3" AND FORMULA(Issue![Balance]3) = "" AND FORMULA(Issue![Balance]4) = "=RC2-RC3" AND FORMULA(Issue![Balance]9) = "=RC2-RC3"
TEST     fixture=issue_01.xlsx call=FillBalanceFormulas(Issue) assert=RESULT = 7 AND COUNTA(Issue![Batch No]) = PRE_COUNTA(Issue![Batch No]) AND SUM(Issue![Opening Stock]) = PRE_SUM(Issue![Opening Stock]) AND SUM(Issue![Issued]) = PRE_SUM(Issue![Issued])
TEST     fixture=issue_01.xlsx call=FillBalanceFormulas(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Notes![Note]) = PRE_COUNTA(Notes![Note])
