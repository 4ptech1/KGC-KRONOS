VCL 0.6
TASK
# ASSUME: Weld Log headers in row 1 include "Weld Seam No" and "Remark"
# ASSUME: seam numbers are trimmed and compared ignoring case, so 'ws-12' and ' WS-12 ' are one seam
# ASSUME: a row whose seam number is blank is skipped: it is never a repeat and its Remark is left alone
# ASSUME: a repeated seam gets Remark Repeat, replacing any text there; every other row with a seam has its Remark cleared
# ASSUME: the real sheet is "Weld Log"; the TESTs use WeldLog
GOAL     Mark every Weld Log row whose seam number appears more than once with Repeat in Remark, clear Remark on the other rows and return how many rows were marked.
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
OP         DICT_ADD(counts, TRIM(CELL($ws, r, cSeam)), 1, mode=sum)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cSeam) var=q
OP         SKIP_IF IS_EMPTY(CELL($ws, q, cSeam))
OP         DICT_GET(counts, TRIM(CELL($ws, q, cSeam)), into=seen)
OP         IF seen > 1
OP           SET_VALUE(CELL($ws, q, cRemark), "Repeat")
OP           LET(marked, marked + 1)
OP         ELSE
OP           CLEAR(CELL($ws, q, cRemark), what=contents)
OP       RETURN(marked)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  weld_log_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Welder ID,C:Weld Date,D:Process,E:Visual,F:Remark rows=10 data=A:values("WS-03"x1," ws-12 "x1,"WS-07"x1,"WS-12"x2,"WS-03 "x1,"WS-21"x1,""x1,"WS-30"x1,"WS-31"x1)+boundary("WS-03"," ws-12 ","WS-07","WS-12","WS-03 ","WS-21","","WS-12","WS-30","WS-31");B:text(5);C:dates(2026-01-05,2026-09-28);D:values("GTAW"x5,"SMAW"x5);E:values("Pass"x8,"Fail"x2);F:values(""x4,"Hold"x2,"Repeat"x3,"Check"x1)+boundary("","Hold","Repeat","","Repeat","Hold","Check","","Repeat","")
FIXTURE  weld_log_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=weld_log_01.xlsx call=MarkRepeatedSeams(WeldLog) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTIF(WeldLog![Remark], "Repeat") = 5 AND COUNTIF(WeldLog![Remark], "Hold") = 0 AND COUNTIF(WeldLog![Remark], "Check") = 1 AND COUNTA(WeldLog![Remark]) = 7
TEST     fixture=weld_log_01.xlsx call=MarkRepeatedSeams(WeldLog) assert=VALUE(WeldLog![Remark]2) = "Repeat" AND VALUE(WeldLog![Remark]3) = "Repeat" AND VALUE(WeldLog![Remark]4) = "" AND VALUE(WeldLog![Remark]5) = "Repeat" AND VALUE(WeldLog![Remark]6) = "Repeat" AND VALUE(WeldLog![Remark]7) = "" AND VALUE(WeldLog![Remark]8) = "Check" AND VALUE(WeldLog![Remark]9) = "Repeat" AND VALUE(WeldLog![Remark]10) = "" AND VALUE(WeldLog![Remark]11) = ""
TEST     fixture=weld_log_01.xlsx call=MarkRepeatedSeams(WeldLog) assert=ERROR_COUNT = 0 AND COUNTA(WeldLog![Weld Seam No]) = PRE_COUNTA(WeldLog![Weld Seam No]) AND COUNTA(WeldLog![Welder ID]) = PRE_COUNTA(WeldLog![Welder ID]) AND COUNTA(WeldLog![Process]) = PRE_COUNTA(WeldLog![Process])
TEST     fixture=weld_log_01.xlsx call=MarkRepeatedSeams(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(WeldLog![Remark]) = PRE_COUNTA(WeldLog![Remark])
