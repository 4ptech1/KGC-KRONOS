VCL 0.6
TASK
# ASSUME: Weld Log headers in row 1 include "Welder ID"
# ASSUME: the Welder ID is matched ignoring case and stray spaces, on the argument too; the other rows are hidden, never deleted
# ASSUME: the count returned is the number of data rows left visible; blank Welder ID rows are hidden unless the argument is blank too
GOAL     Show only the Weld Log rows of the given Welder ID, hide all the others and return how many rows are visible below the headings.
PROC     ShowWelderRows
KIND     Function
MODULE   modWeldLog
ARGS     ws As Worksheet, welder As String
RET      Long
VAR      cWelder As Long
VAR      shown As Long
PRE      COL($ws, "Welder ID") > 0
MODIFIES $ws
OP       LET(cWelder, COL($ws, "Welder ID"))
OP       FILTER(USED_RANGE($ws), col=cWelder, crit=$welder)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cWelder) var=r
OP         IF SAME(CELL($ws, r, cWelder), $welder)
OP           LET(shown, shown + 1)
OP       RETURN(shown)
PERF     screen_off
ERR      standard
FIXTURE  weld_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Welder ID,C:Visual rows=9 data=A:seq(1,1);B:values("W-01"x3,"  w-01"x2,"W-02"x2,"W-011"x1,""x1)+boundary("W-01","  w-01","W-011","W-02","");C:values("Accept"x6,"Reject"x3)
FIXTURE  weld_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=weld_01.xlsx call=ShowWelderRows(WeldLog, "W-01") runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND VISIBLE_COUNT(WeldLog![Welder ID]) = 6 AND VISIBLE_COUNT(WeldLog![Weld Seam No]) = 6
TEST     fixture=weld_01.xlsx call=ShowWelderRows(WeldLog, " W-02 ") assert=RESULT = 2 AND VISIBLE_COUNT(WeldLog![Welder ID]) = 3 AND COUNTA(WeldLog![Weld Seam No]) = PRE_COUNTA(WeldLog![Weld Seam No]) AND COUNTA(WeldLog![Visual]) = PRE_COUNTA(WeldLog![Visual])
TEST     fixture=weld_01.xlsx call=ShowWelderRows(WeldLog, "W-99") assert=RESULT = 0 AND VISIBLE_COUNT(WeldLog![Weld Seam No]) = 1 AND ERROR_COUNT = 0
TEST     fixture=weld_01.xlsx call=ShowWelderRows(Notes, "W-01") assert=RESULT = -1 AND ERROR_COUNT = 1
