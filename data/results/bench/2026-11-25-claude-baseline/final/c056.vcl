VCL 0.7
TASK
# ASSUME: Weld Log headers in row 1 include "Weld Seam No" and "Remark"
# ASSUME: seam numbers are compared after CLEAN_TEXT and ignoring case, so 'ws-12' and ' WS-12 ' are the same seam
# ASSUME: rows with a blank Weld Seam No are never marked, and their Remark is cleared
GOAL     Write Repeat in Remark for every weld whose seam number appears more than once, clear Remark on the others and return how many rows were marked.
PROC     MarkRepeatedSeams
KIND     Function
MODULE   modWeldLog
ARGS     ws As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cRemark As Long
VAR      counts As Dictionary
VAR      seen As Long
VAR      marked As Long
PRE      COL($ws, "Weld Seam No") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Remark"]
OP       LET(cSeam, COL($ws, "Weld Seam No"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cSeam))
OP         DICT_ADD(counts, CLEAN_TEXT(CELL($ws, r, cSeam)), 1, mode=sum)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cSeam) var=r
OP         IF IS_EMPTY(CELL($ws, r, cSeam))
OP           SET_VALUE(CELL($ws, r, cRemark), "")
OP         ELSE
OP           DICT_GET(counts, CLEAN_TEXT(CELL($ws, r, cSeam)), into=seen)
OP           IF seen > 1
OP             SET_VALUE(CELL($ws, r, cRemark), "Repeat")
OP             LET(marked, marked + 1)
OP           ELSE
OP             SET_VALUE(CELL($ws, r, cRemark), "")
OP       RETURN(marked)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  weld_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Welder ID,C:Weld Date,D:Process,E:Visual,F:Remark rows=8 data=A:values("WS-1"x1,"ws-1"x1," WS-3 "x1,"WS-2"x1,"WS-3"x1,"WS-4"x1,""x1,"WS-5"x1)+boundary("WS-1","ws-1"," WS-3 ","WS-2","WS-3","WS-4","","WS-5");B:text(4);C:dates(2026-01-01,2026-09-30);D:values("GTAW"x4,"SMAW"x4);E:values("Pass"x8);F:values("old"x8)
FIXTURE  weld_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=weld_01.xlsx call=MarkRepeatedSeams(WeldLog) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(WeldLog![Remark], "Repeat") = 4 AND COUNTIF(WeldLog![Remark], "old") = 0
TEST     fixture=weld_01.xlsx call=MarkRepeatedSeams(WeldLog) assert=COUNTA(WeldLog![Weld Seam No]) = PRE_COUNTA(WeldLog![Weld Seam No]) AND VALUE(WeldLog![Remark]2) = "Repeat" AND VALUE(WeldLog![Remark]3) = "Repeat" AND VALUE(WeldLog![Remark]5) = "" AND VALUE(WeldLog![Remark]6) = "Repeat"
TEST     fixture=weld_01.xlsx call=MarkRepeatedSeams(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
