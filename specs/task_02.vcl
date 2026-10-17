VCL 0.1
TASK
# ASSUME: Material Traceability headers (row 1) include "Part No", "Plate/Heat No" and "MTC No"; column order does not matter
# ASSUME: rows with a blank Part No are skipped
# ASSUME: the macro owns the fill colour of listed rows: it clears the fill first, so fixed rows lose their red
# ASSUME: the whole row is highlighted, from column A to the last header column
PROC     FlagMissingTraceability
KIND     Function
MODULE   modMaterials
ARGS     ws As Worksheet
RET      Long
VAR      cPart As Long
VAR      cHeat As Long
VAR      cMtc As Long
VAR      lastCol As Long
VAR      flagged As Long
PRE      COL($ws, "Part No") > 0; COL($ws, "Plate/Heat No") > 0; COL($ws, "MTC No") > 0
OP       LET(cPart, COL($ws, "Part No"))
OP       LET(cHeat, COL($ws, "Plate/Heat No"))
OP       LET(cMtc, COL($ws, "MTC No"))
OP       LET(lastCol, LASTCOL($ws, 1))
OP       LET(flagged, 0)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cPart) var=r
OP         IF NOT IS_EMPTY(CELL($ws, r, cPart))
OP           FORMAT(ROW_RANGE($ws, r, 1, lastCol), fill=none)
OP           IF IS_EMPTY(CELL($ws, r, cHeat)) OR IS_EMPTY(CELL($ws, r, cMtc))
OP             FORMAT(ROW_RANGE($ws, r, 1, lastCol), fill=#FF0000)
OP             LET(flagged, flagged + 1)
OP       RETURN(flagged)
PERF     screen_off
ERR      standard
FIXTURE  traceability_01.xlsx sheet=Traceability cols=A:Part No,B:Plate/Heat No,C:MTC No,D:Remarks rows=40 stale=6 data=A:seq(1,1);B:text(8)+boundary("")+blank(4);C:text(10)+blank(3);D:text(10)+boundary("")
FIXTURE  traceability_01.xlsx sheet=Notes cols=A:Note rows=3 data=A:text(20)
TEST     fixture=traceability_01.xlsx call=FlagMissingTraceability(Traceability) runs=2 assert=RESULT >= 5 AND RESULT <= 8 AND FILL_COUNT(Traceability![Part No], #FF0000) = RESULT AND FILL_COUNT(Traceability![Part No]2, #FF0000) = 1 AND ERROR_COUNT = 0
TEST     fixture=traceability_01.xlsx call=FlagMissingTraceability(Traceability) assert=COUNTA(Traceability![Plate/Heat No]) = PRE_COUNTA(Traceability![Plate/Heat No]) AND COUNTA(Traceability![MTC No]) = PRE_COUNTA(Traceability![MTC No])
TEST     fixture=traceability_01.xlsx call=FlagMissingTraceability(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1 AND FILL_COUNT(Notes![Note], #FF0000) = 0
