VCL 0.6
TASK
# ASSUME: GRN headers in row 1 include "Inspector"; Ali is matched ignoring case and extra spaces
# ASSUME: a row whose Inspector is blank or anyone else (Ali Khan too) is hidden; no row is deleted
# ASSUME: the number of visible GRN rows, header not counted, is returned as the result, not shown in a message
# ASSUME: each run re-applies the filter; VCL can't read whether a row is hidden or show every row, so a second run does not show the other rows again
GOAL     Hide every row of the GRN sheet whose Inspector is not Ali and return how many GRN rows are still visible.
PROC     HideNonAliGrnRows
KIND     Function
MODULE   modGrn
ARGS     ws As Worksheet
RET      Long
VAR      cInsp As Long
VAR      n As Long
PRE      COL($ws, "Inspector") > 0
MODIFIES $ws
OP       LET(cInsp, COL($ws, "Inspector"))
OP       FILTER(USED_RANGE($ws), col=cInsp, crit="Ali")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cInsp) var=r
OP         IF SAME(CELL($ws, r, cInsp), "Ali") THEN LET(n, n + 1)
OP       RETURN(n)
ERR      standard
FIXTURE  grn_01.xlsx sheet=GRN cols=A:GRN No,B:Part No,C:Received Qty,D:Inspector rows=10 data=A:seq(5001,1);B:text(8);C:int(1,200);D:values(" ali "x1,"Sara"x2,"Ali"x4,""x1,"Ali Khan"x1,"ALI"x1)+boundary(" ali ","Sara","Ali","","Ali Khan","ALI","Ali","Sara","Ali","Ali")
FIXTURE  grn_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=grn_01.xlsx call=HideNonAliGrnRows(GRN) runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND VISIBLE_COUNT(GRN![GRN No]) = 7 AND VISIBLE_COUNT(GRN![Inspector]) = 7 AND VISIBLE_COUNT(GRN![GRN No]2) = 1 AND VISIBLE_COUNT(GRN![GRN No]5) = 0 AND VISIBLE_COUNT(GRN![GRN No]6) = 0 AND VISIBLE_COUNT(GRN![GRN No]7) = 1
TEST     fixture=grn_01.xlsx call=HideNonAliGrnRows(GRN) assert=RESULT = VISIBLE_COUNT(GRN![GRN No]) - 1 AND ERROR_COUNT = 0 AND COUNTA(GRN![GRN No]) = PRE_COUNTA(GRN![GRN No]) AND COUNTA(GRN![Part No]) = PRE_COUNTA(GRN![Part No]) AND SUM(GRN![Received Qty]) = PRE_SUM(GRN![Received Qty]) AND COUNTA(GRN![Inspector]) = PRE_COUNTA(GRN![Inspector])
TEST     fixture=grn_01.xlsx call=HideNonAliGrnRows(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
