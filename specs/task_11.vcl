VCL 0.2
TASK
# ASSUME: WeldLog headers are in row 1 and include "Weld Seam No" and "Visual"; data starts in row 2
# ASSUME: "Accepted" is matched ignoring case and extra spaces; blank, error or any other Visual value counts as not Accepted
# ASSUME: a Weld Seam No made only of spaces counts as blank
# ASSUME: the macro owns the row fills on WeldLog: every data row is reset to no fill, then flagged rows get red #FF0000 across columns 1 to the last header column
# ASSUME: rows with a blank Weld Seam No are never highlighted and are left with no fill
# ASSUME: the last data row is the lowest row holding a value in any header column
PROC     HighlightRejectedWelds
KIND     Function
MODULE   modWeldCheck
ARGS     ws As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cVisual As Long
VAR      lastCol As Long
VAR      lastRow As Long
VAR      n As Long
VAR      flagged As Long
PRE      COL($ws, "Weld Seam No") > 0; COL($ws, "Visual") > 0
MODIFIES $ws fills
OP       LET(cSeam, COL($ws, "Weld Seam No"))
OP       LET(cVisual, COL($ws, "Visual"))
OP       LET(lastCol, LASTCOL($ws, 1))
OP       LET(lastRow, 1)
OP       LOOP_COLS dir=right from=1 to=lastCol var=c
OP         LET(n, LASTROW($ws, c))
OP         IF n > lastRow THEN LET(lastRow, n)
OP       LET(flagged, 0)
OP       LOOP_ROWS dir=down from=2 to=lastRow var=r
OP         IF IS_EMPTY(CELL($ws, r, cSeam))
OP           FORMAT(ROW_RANGE($ws, r, 1, lastCol), fill=none)
OP         ELSE
OP           IF SAME(CELL($ws, r, cVisual), "Accepted")
OP             FORMAT(ROW_RANGE($ws, r, 1, lastCol), fill=none)
OP           ELSE
OP             FORMAT(ROW_RANGE($ws, r, 1, lastCol), fill=red)
OP             LET(flagged, flagged + 1)
OP       RETURN(flagged)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  weldlog_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Joint Type,C:Welder ID,D:Visual,E:Remarks rows=30 stale=10 data=A:seq(1001,1)+boundary(1001,"")+blank(2);B:values("Butt"x15,"Fillet"x15);C:values("W-01"x10,"W-02"x10,"W-03"x10);D:values("Accepted"x14," accepted "x2,"ACCEPTED"x2,"Rejected"x6,"Repair"x6)+boundary("Rejected","Repair")+blank(2);E:text(15)+blank(3)
FIXTURE  weldlog_01.xlsx sheet=Notes cols=A:Note rows=3 data=A:text(10)
FIXTURE  weldlog_02.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Joint Type,C:Welder ID,D:Visual,E:Remarks rows=30 stale=30 data=A:seq(1001,1)+boundary(1001,"")+blank(2);B:values("Butt"x15,"Fillet"x15);C:values("W-01"x10,"W-02"x10,"W-03"x10);D:values("Accepted"x14," accepted "x2,"ACCEPTED"x2,"Rejected"x6,"Repair"x6)+boundary("Rejected","Repair")+blank(2);E:text(15)+blank(3)
TEST     fixture=weldlog_01.xlsx call=HighlightRejectedWelds(WeldLog) runs=2 assert=FILL_COUNT(WeldLog![Visual], #FF0000) = RESULT AND RESULT = COUNTIFS(WeldLog![Weld Seam No], "<>", WeldLog![Visual], "Rejected") + COUNTIFS(WeldLog![Weld Seam No], "<>", WeldLog![Visual], "Repair") + COUNTIFS(WeldLog![Weld Seam No], "<>", WeldLog![Visual], "=") AND ERROR_COUNT = 0
TEST     fixture=weldlog_01.xlsx call=HighlightRejectedWelds(WeldLog) runs=2 assert=FILL_COUNT(WeldLog![Weld Seam No]2, #FF0000) = 1 AND FILL_COUNT(WeldLog![Visual]2, #FF0000) = 1 AND FILL_COUNT(WeldLog![Remarks]2, #FF0000) = 1 AND FILL_COUNT(WeldLog![Visual]3, #FF0000) = 0 AND ERROR_COUNT = 0
TEST     fixture=weldlog_02.xlsx call=HighlightRejectedWelds(WeldLog) runs=2 assert=FILL_COUNT(WeldLog![Visual], #FF0000) = RESULT AND RESULT = COUNTIFS(WeldLog![Weld Seam No], "<>", WeldLog![Visual], "Rejected") + COUNTIFS(WeldLog![Weld Seam No], "<>", WeldLog![Visual], "Repair") + COUNTIFS(WeldLog![Weld Seam No], "<>", WeldLog![Visual], "=") AND FILL_COUNT(WeldLog![Weld Seam No]3, #FF0000) = 0 AND FILL_COUNT(WeldLog![Visual]3, #FF0000) = 0 AND FILL_COUNT(WeldLog![Remarks]3, #FF0000) = 0 AND ERROR_COUNT = 0
TEST     fixture=weldlog_01.xlsx call=HighlightRejectedWelds(WeldLog) assert=FILL_COUNT(WeldLog![Weld Seam No], #FF0000) = RESULT AND COUNTA(WeldLog![Visual]) = PRE_COUNTA(WeldLog![Visual]) AND COUNTIF(WeldLog![Visual], "Accepted") = PRE_COUNTIF(WeldLog![Visual], "Accepted") AND COUNTA(WeldLog![Weld Seam No]) = PRE_COUNTA(WeldLog![Weld Seam No]) AND ERROR_COUNT = 0
TEST     fixture=weldlog_01.xlsx call=HighlightRejectedWelds(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
