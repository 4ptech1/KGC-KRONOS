VCL 0.6
TASK
# ASSUME: Weld log headers in row 1 include "Welder ID" and "Visual"
# ASSUME: the real sheets are "Weld Log" and "Welds Per Welder"; the TESTs use WeldLog and PerWelder
# ASSUME: Welds Per Welder already exists and is rebuilt on every run, with the headers Welder ID, Welds and Rejected
# ASSUME: Welder ID is trimmed and matched ignoring case; a row with a blank Welder ID is not counted
# ASSUME: a weld is Rejected when Visual says Rejected, ignoring case and extra spaces
# ASSUME: welders with the same number of welds may appear in either order
GOAL     Count each welder's welds and rejected welds from the weld log and list them on the Welds Per Welder sheet, most welds first.
PROC     CountWeldsPerWelder
KIND     Function
MODULE   modWelds
ARGS     weldLog As Worksheet, out As Worksheet
RET      Long
VAR      cWelder As Long
VAR      cVisual As Long
VAR      welds As Dictionary
VAR      rejects As Dictionary
VAR      wCount As Long
VAR      rCount As Long
VAR      outRow As Long
PRE      COL($weldLog, "Welder ID") > 0; COL($weldLog, "Visual") > 0
MODIFIES $out
OP       LET(cWelder, COL($weldLog, "Welder ID"))
OP       LET(cVisual, COL($weldLog, "Visual"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Welder ID", "Welds", "Rejected")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($weldLog, cWelder) var=r
OP         SKIP_IF IS_EMPTY(CELL($weldLog, r, cWelder))
OP         DICT_ADD(welds, TRIM(CELL($weldLog, r, cWelder)), 1, mode=sum)
OP         DICT_ADD(rejects, TRIM(CELL($weldLog, r, cWelder)), IIF(SAME(CELL($weldLog, r, cVisual), "Rejected"), 1, 0), mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(welds)
OP         DICT_GET(welds, k, into=wCount)
OP         DICT_GET(rejects, k, into=rCount)
OP         WRITE_ROW($out, outRow, k, wCount, rCount)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=2, order=desc, header=yes)
OP       RETURN(COUNT(welds))
PERF     screen_off
ERR      standard
FIXTURE  weld_log_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Welder ID,C:Joint Type,D:Weld Date,E:Visual rows=14 data=A:seq(1,1);B:values("W-05"x1,"W-09"x2,"W-03"x4,"W-12"x5," W-12 "x1,""x1)+boundary("W-05","W-09","W-03","W-12","W-12","","W-03","W-12","W-09","W-03"," W-12 ","W-03","W-12","W-12");C:values("Butt"x8,"Fillet"x6);D:dates(2026-01-01,2026-06-30);E:values("Accepted"x8,"Rejected"x5," rejected "x1)+boundary("Accepted","Rejected","Accepted","Rejected","Accepted","Rejected","Rejected","Accepted","Accepted","Accepted","Rejected"," rejected ","Accepted","Accepted")
FIXTURE  weld_log_01.xlsx sheet=PerWelder cols=A:Welder ID,B:Welds,C:Rejected rows=0 stale=2
TEST     fixture=weld_log_01.xlsx call=CountWeldsPerWelder(WeldLog, PerWelder) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND ROWS_USED(PerWelder) = 5 AND COUNTIF(PerWelder![Welder ID], "stale") = 0 AND VALUE(PerWelder![Welder ID]2) = "W-12" AND VALUE(PerWelder![Welds]2) = 6 AND VALUE(PerWelder![Rejected]2) = 2 AND VALUE(PerWelder![Welder ID]3) = "W-03" AND VALUE(PerWelder![Welds]3) = 4 AND VALUE(PerWelder![Rejected]3) = 2
TEST     fixture=weld_log_01.xlsx call=CountWeldsPerWelder(WeldLog, PerWelder) assert=RESULT = 4 AND ERROR_COUNT = 0 AND VALUE(PerWelder![Welder ID]4) = "W-09" AND VALUE(PerWelder![Welds]4) = 2 AND VALUE(PerWelder![Rejected]4) = 1 AND VALUE(PerWelder![Welder ID]5) = "W-05" AND VALUE(PerWelder![Welds]5) = 1 AND VALUE(PerWelder![Rejected]5) = 0
TEST     fixture=weld_log_01.xlsx call=CountWeldsPerWelder(WeldLog, PerWelder) assert=SUM(PerWelder![Welds]) = COUNTA(WeldLog![Welder ID]) - 1 AND SUM(PerWelder![Rejected]) = 5 AND COUNTA(WeldLog![Welder ID]) = PRE_COUNTA(WeldLog![Welder ID]) AND COUNTA(WeldLog![Visual]) = PRE_COUNTA(WeldLog![Visual]) AND COUNTIF(WeldLog![Visual], "Rejected") = PRE_COUNTIF(WeldLog![Visual], "Rejected")
TEST     fixture=weld_log_01.xlsx call=CountWeldsPerWelder(PerWelder, WeldLog) assert=RESULT = -1 AND ERROR_COUNT = 1
