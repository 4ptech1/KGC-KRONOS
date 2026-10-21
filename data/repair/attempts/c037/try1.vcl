VCL 0.6
TASK
# ASSUME: WeldLog headers in row 1 include "Weld Seam No" and "Joint Type"
# ASSUME: each row with a Joint Type is one weld; rows with a blank Joint Type are ignored
# ASSUME: joint types are matched ignoring case and extra spaces; the first spelling seen is written
# ASSUME: the real sheets are "Weld Log" and "Joint Summary"; the summary sheet exists and is rebuilt on every run
# ASSUME: the summary has the headings "Joint Type" and "Weld Count"; joint types with equal counts come in no fixed order
GOAL     Count the welds per joint type from the weld log and write them on the Joint Summary sheet with the biggest count first.
PROC     CountWeldsByJointType
KIND     Function
MODULE   modWeldSummary
ARGS     log As Worksheet, out As Worksheet
RET      Long
VAR      cJoint As Long
VAR      counts As Dictionary
VAR      got As Long
VAR      outRow As Long
PRE      COL($log, "Weld Seam No") > 0; COL($log, "Joint Type") > 0
MODIFIES $out
OP       LET(cJoint, COL($log, "Joint Type"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Joint Type", "Weld Count")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($log, cJoint) var=r
OP         SKIP_IF IS_EMPTY(CELL($log, r, cJoint))
OP         DICT_ADD(counts, CLEAN_TEXT(CELL($log, r, cJoint)), 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         WRITE_ROW($out, outRow, k, got)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=2, order=desc, header=yes)
OP       RETURN(COUNT(counts))
PERF     screen_off calc_manual
ERR      standard
FIXTURE  weldlog_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Welder ID,C:Weld Date,D:Joint Type rows=15 data=A:seq(1,1);B:text(4);C:dates(2026-01-01,2026-09-30);D:values("Butt"x5," BUTT "x1,"Fillet"x4,"Corner"x2,"Lap"x1,""x2)+boundary("Butt","","Fillet"," BUTT ")
FIXTURE  weldlog_01.xlsx sheet=JointSummary cols=A:Joint Type,B:Weld Count rows=0 stale=2
TEST     fixture=weldlog_01.xlsx call=CountWeldsByJointType(WeldLog, JointSummary) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND ROWS_USED(JointSummary) = 5 AND COUNTIF(JointSummary![Joint Type], "stale") = 0 AND VALUE(JointSummary![Joint Type]2) = "Butt" AND VALUE(JointSummary![Weld Count]2) = 6 AND VALUE(JointSummary![Weld Count]3) = 4 AND VALUE(JointSummary![Weld Count]4) = 2 AND VALUE(JointSummary![Weld Count]5) = 1
TEST     fixture=weldlog_01.xlsx call=CountWeldsByJointType(WeldLog, JointSummary) assert=SUM(JointSummary![Weld Count]) = PRE_COUNTA(WeldLog![Joint Type]) - 1 AND COUNTA(WeldLog![Joint Type]) = PRE_COUNTA(WeldLog![Joint Type]) AND COUNTA(WeldLog![Weld Seam No]) = PRE_COUNTA(WeldLog![Weld Seam No]) AND ERROR_COUNT = 0
TEST     fixture=weldlog_01.xlsx call=CountWeldsByJointType(JointSummary, WeldLog) assert=RESULT = -1 AND ERROR_COUNT = 1
