VCL 0.6
TASK
# ASSUME: Weld Log headers in row 1 include "Weld Seam No", "Welder ID" and "Visual"
# ASSUME: a cell is changed only when its cleaned text (trimmed, upper case, inner space runs made one space) differs from what it holds
# ASSUME: blank cells and error cells are left alone; a cell is counted once, and a cell that holds a number but is already clean is not touched
# ASSUME: the Visual column is never touched
GOAL     Clean every Weld Seam No and Welder ID in the Weld Log to trimmed upper case with single inner spaces, and return how many cells changed.
PROC     CleanWeldLogIds
KIND     Function
MODULE   modWeldLog
ARGS     ws As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cWelder As Long
VAR      changed As Long
VAR      clean As String
PRE      COL($ws, "Weld Seam No") > 0; COL($ws, "Welder ID") > 0; COL($ws, "Visual") > 0
MODIFIES $ws["Weld Seam No", "Welder ID"]
OP       LET(cSeam, COL($ws, "Weld Seam No"))
OP       LET(cWelder, COL($ws, "Welder ID"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cSeam) var=r
OP         IF NOT IS_EMPTY(CELL($ws, r, cSeam))
OP           IF NOT IS_ERROR(CELL($ws, r, cSeam))
OP             LET(clean, UPPER(CLEAN_TEXT(CELL($ws, r, cSeam))))
OP             IF clean <> CELL($ws, r, cSeam)
OP               SET_VALUE(CELL($ws, r, cSeam), clean)
OP               LET(changed, changed + 1)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cWelder) var=r
OP         IF NOT IS_EMPTY(CELL($ws, r, cWelder))
OP           IF NOT IS_ERROR(CELL($ws, r, cWelder))
OP             LET(clean, UPPER(CLEAN_TEXT(CELL($ws, r, cWelder))))
OP             IF clean <> CELL($ws, r, cWelder)
OP               SET_VALUE(CELL($ws, r, cWelder), clean)
OP               LET(changed, changed + 1)
OP       RETURN(changed)
PERF     screen_off
ERR      standard
FIXTURE  weldlog_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Welder ID,C:Visual rows=8 data=A:values("w-1"x1,"w-2"x1,"  W-3 "x1,"w  4"x1,""x1,"W-6"x1,"  "x1,"w - 8"x1)+boundary("w-1","w-2","  W-3 ","w  4","","W-6","  ","w - 8");B:values("ab1"x1,"AB2"x1,"  ab 3"x1,""x1,"AB5"x1,"ab  6  x"x1,"AB7"x1,"ab8"x1)+boundary("ab1","AB2","  ab 3","","AB5","ab  6  x","AB7","ab8");C:values("pass"x2,"Fail "x2,"  ok"x2,""x2)+boundary("pass","Fail ","  ok","","pass","Fail ","  ok","")
FIXTURE  weldlog_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=weldlog_01.xlsx call=CleanWeldLogIds(WeldLog) runs=2 assert=ERROR_COUNT = 0 AND VALUE(WeldLog![Weld Seam No]2) = "W-1" AND VALUE(WeldLog![Weld Seam No]3) = "W-2" AND VALUE(WeldLog![Weld Seam No]4) = "W-3" AND VALUE(WeldLog![Weld Seam No]5) = "W 4" AND VALUE(WeldLog![Weld Seam No]9) = "W - 8" AND VALUE(WeldLog![Welder ID]4) = "AB 3" AND VALUE(WeldLog![Welder ID]7) = "AB 6 X"
TEST     fixture=weldlog_01.xlsx call=CleanWeldLogIds(WeldLog) assert=RESULT = 9 AND ERROR_COUNT = 0 AND COUNTA(WeldLog![Visual]) = PRE_COUNTA(WeldLog![Visual]) AND COUNTA(WeldLog![Weld Seam No]) = PRE_COUNTA(WeldLog![Weld Seam No]) AND COUNTA(WeldLog![Welder ID]) = PRE_COUNTA(WeldLog![Welder ID])
TEST     fixture=weldlog_01.xlsx call=CleanWeldLogIds(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
