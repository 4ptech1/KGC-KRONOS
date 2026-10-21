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
VAR      hasOther As Boolean
VAR      removed As Long
PRE      COL($ws, "Weld Seam No") > 0; COL($ws, "Welder ID") > 0
MODIFIES $ws
OP       LET(cSeam, COL($ws, "Weld Seam No"))
OP       LET(cWelder, COL($ws, "Welder ID"))
OP       LET(lastCol, LASTCOL($ws, 1))
OP       LOOP_COLS dir=right from=1 to=lastCol var=k
OP         IF LASTROW($ws, k) > last THEN LET(last, LASTROW($ws, k))
OP       LOOP_ROWS dir=up from=last to=2 var=r
OP         SKIP_IF NOT IS_EMPTY(CELL($ws, r, cSeam))
OP         SKIP_IF NOT IS_EMPTY(CELL($ws, r, cWelder))
OP         LET(hasOther, FALSE)
OP         LOOP_COLS dir=right from=1 to=lastCol var=c
OP           IF c <> cSeam AND c <> cWelder AND NOT IS_EMPTY(CELL($ws, r, c)) THEN LET(hasOther, TRUE)
OP         IF NOT hasOther
OP           DELETE_ROW($ws, r)
OP           LET(removed, removed + 1)
OP       RETURN(removed)
ERR      standard
FIXTURE  welds_01.xlsx sheet=Welds cols=A:Item No,B:Weld Seam No,C:Welder ID,D:Joint Type,E:Remarks rows=10 data=A:text(4)+boundary("I-1","","","","","","I-8","","","");B:text(4)+boundary("S-01","","S-03","","","  ","","","","");C:text(4)+boundary("W-1","","","W-4","","","","","","");D:text(4)+boundary("BW","","","","","","","NZ","","");E:text(4)+boundary("","","","","","","","","R-10","  ")
FIXTURE  welds_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  welds_02.xlsx sheet=Welds cols=A:Item No,B:Weld Seam No,C:Welder ID,D:Joint Type,E:Remarks rows=4 data=A:text(4)+boundary("","I-1","","I-5");B:text(4)+boundary("","S-01","","");C:text(4)+boundary("","W-1","","");D:text(4)+boundary("","BW","","");E:text(4)+boundary("","R-1","","")
TEST     fixture=welds_01.xlsx call=RemoveEmptyWeldRows(Welds) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND ROWS_USED(Welds) = 7 AND COUNTA(Welds![Weld Seam No]) = 3 AND COUNTA(Welds![Welder ID]) = 3
TEST     fixture=welds_01.xlsx call=RemoveEmptyWeldRows(Welds) assert=RESULT = 4 AND ERROR_COUNT = 0 AND ROWS_USED(Welds) = PRE_ROWS_USED(Welds) - RESULT AND COUNTA(Welds![Item No]) = PRE_COUNTA(Welds![Item No]) AND COUNTA(Welds![Joint Type]) = PRE_COUNTA(Welds![Joint Type])
TEST     fixture=welds_01.xlsx call=RemoveEmptyWeldRows(Welds) assert=VALUE(Welds![Weld Seam No]3) = "S-03" AND VALUE(Welds![Welder ID]4) = "W-4" AND COUNTA(Welds![Welder ID]) = 3 AND VALUE(Welds![Item No]5) = "I-8" AND VALUE(Welds![Joint Type]6) = "NZ" AND VALUE(Welds![Remarks]7) = "R-10"
TEST     fixture=welds_02.xlsx call=RemoveEmptyWeldRows(Welds) assert=RESULT = 2 AND ERROR_COUNT = 0 AND ROWS_USED(Welds) = 3 AND VALUE(Welds![Item No]3) = "I-5"
TEST     fixture=welds_01.xlsx call=RemoveEmptyWeldRows(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
