VCL 0.6
TASK
# ASSUME: Weld log headers in row 1 include "Weld Seam No" and "Welder ID"; every column of the log has a header
# ASSUME: a cell holding only spaces counts as blank
# ASSUME: leftover empty rows are rows where Weld Seam No and Welder ID are blank and every other cell is blank too
# ASSUME: a row with both IDs blank but other data in it is kept, because deleting it would destroy data
GOAL     Delete the leftover empty rows of the weld log, those with both Weld Seam No and Welder ID blank, keep any row holding data, and return how many rows were deleted.
PROC     RemoveEmptyWeldRows
KIND     Function
MODULE   modWeldCleanup
ARGS     ws As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cWelder As Long
VAR      lastCol As Long
VAR      last As Long
VAR      filled As Long
VAR      removed As Long
PRE      COL($ws, "Weld Seam No") > 0; COL($ws, "Welder ID") > 0
MODIFIES $ws
OP       LET(cSeam, COL($ws, "Weld Seam No"))
OP       LET(cWelder, COL($ws, "Welder ID"))
OP       LET(lastCol, LASTCOL($ws, 1))
OP       LET(last, 1)
OP       LOOP_COLS dir=right from=1 to=lastCol var=k
OP         IF LASTROW($ws, k) > last THEN LET(last, LASTROW($ws, k))
OP       LOOP_ROWS dir=up from=last to=2 var=r
OP         SKIP_IF NOT IS_EMPTY(CELL($ws, r, cSeam))
OP         SKIP_IF NOT IS_EMPTY(CELL($ws, r, cWelder))
OP         LET(filled, 0)
OP         LOOP_COLS dir=right from=1 to=lastCol var=c
OP           IF NOT IS_EMPTY(CELL($ws, r, c)) THEN LET(filled, filled + 1)
OP         IF filled = 0
OP           DELETE_ROW($ws, r)
OP           LET(removed, removed + 1)
OP       RETURN(removed)
ERR      standard
FIXTURE  welds_01.xlsx sheet=Welds cols=A:Weld Seam No,B:Welder ID,C:Joint Type,D:Weld Date rows=10 data=A:text(6)+boundary("S-01","","S-03","","","  ","","","S-09","");B:text(4)+boundary("W-1","","","W-4","","","","","W-9","");C:text(4)+boundary("BW","","BW","FW","","","NZ","","BW","FW");D:text(10)+boundary("2026-03-01","","2026-03-03","2026-03-04","","","","2026-03-08","2026-03-09","")
FIXTURE  welds_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=welds_01.xlsx call=RemoveEmptyWeldRows(Welds) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND ROWS_USED(Welds) = 8 AND COUNTA(Welds![Weld Seam No]) = 4 AND COUNTA(Welds![Welder ID]) = 4
TEST     fixture=welds_01.xlsx call=RemoveEmptyWeldRows(Welds) assert=RESULT = 3 AND ERROR_COUNT = 0 AND ROWS_USED(Welds) = PRE_ROWS_USED(Welds) - RESULT AND COUNTA(Welds![Joint Type]) = PRE_COUNTA(Welds![Joint Type]) AND COUNTA(Welds![Weld Date]) = PRE_COUNTA(Welds![Weld Date])
TEST     fixture=welds_01.xlsx call=RemoveEmptyWeldRows(Welds) assert=VALUE(Welds![Weld Seam No]2) = "S-01" AND VALUE(Welds![Weld Seam No]3) = "S-03" AND VALUE(Welds![Welder ID]4) = "W-4" AND VALUE(Welds![Joint Type]5) = "NZ" AND VALUE(Welds![Weld Date]6) = "2026-03-08" AND VALUE(Welds![Joint Type]8) = "FW"
TEST     fixture=welds_01.xlsx call=RemoveEmptyWeldRows(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
