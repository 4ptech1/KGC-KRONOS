VCL 0.6
TASK
# ASSUME: the sheets are "Consumable Issue" and "Recent Usage"; Recent Usage exists and is rebuilt on every run
# ASSUME: headers in row 1 of the log include "Issue No", "Issue Date", "Filler Class" and "Qty Issued"
# ASSUME: the last 30 days run from today minus 30 days up to and including today; later dates are left out
# ASSUME: filler classes are trimmed and shown in upper case, so "mig wire" and "MIG Wire" are one class
# ASSUME: rows with a blank class, a non-date Issue Date or a Qty Issued that is not a number are ignored; rows are counted by Issue No
GOAL     Rebuild the Recent Usage sheet with each filler class and its total quantity issued in the last 30 days, biggest total first.
PROC     BuildRecentUsage
KIND     Function
MODULE   modConsumables
ARGS     log As Worksheet, out As Worksheet
RET      Long
VAR      cNo As Long
VAR      cDate As Long
VAR      cClass As Long
VAR      cQty As Long
VAR      sums As Dictionary
VAR      got As Double
VAR      outRow As Long
PRE      COL($log, "Issue No") > 0; COL($log, "Issue Date") > 0; COL($log, "Filler Class") > 0; COL($log, "Qty Issued") > 0
MODIFIES $out
OP       LET(cNo, COL($log, "Issue No"))
OP       LET(cDate, COL($log, "Issue Date"))
OP       LET(cClass, COL($log, "Filler Class"))
OP       LET(cQty, COL($log, "Qty Issued"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Filler Class", "Total Qty Issued")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($log, cNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($log, r, cClass))
OP         SKIP_IF NOT IS_DATE(CELL($log, r, cDate))
OP         SKIP_IF NOT IS_NUMBER(CELL($log, r, cQty))
OP         SKIP_IF CELL($log, r, cDate) < TODAY - 30 OR CELL($log, r, cDate) > TODAY
OP         DICT_ADD(sums, UPPER(TRIM(CELL($log, r, cClass))), CELL($log, r, cQty), mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(sums)
OP         DICT_GET(sums, k, into=got)
OP         WRITE_ROW($out, outRow, k, got)
OP         LET(outRow, outRow + 1)
OP       IF outRow > 3
OP         SORT(USED_RANGE($out), key=2, order=desc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  log_01.xlsx sheet=Log cols=A:Issue No,B:Issue Date,C:Filler Class,D:Qty Issued rows=14 data=A:seq(1,1);B:reldates(-60,10)+boundary(-1,-30,-31,0,1,-10,-5,-2,"",-20,-7,-29,-3,-15);C:values("mig wire"x2,"MIG Wire"x1,"Flux"x3,"flux "x1,"TIG Rod"x2,"tig rod"x1,""x1,"Stick"x2,"Gas"x1)+boundary("mig wire","MIG Wire","Flux","flux ","Flux","TIG Rod","tig rod","","Flux","Stick","TIG Rod","Stick","Gas","mig wire");D:int(0,200)+boundary(10,5,100,8,50,20,"",99,40,3,6,25,0,1)
FIXTURE  log_01.xlsx sheet=Use cols=A:Filler Class,B:Total Qty Issued rows=0 stale=3
FIXTURE  log_02.xlsx sheet=Log cols=A:Issue No,B:Issue Date,C:Filler Class,D:Qty Issued rows=5 data=A:seq(1,1);B:reldates(-5,-1)+boundary(-2,-3,"",-4,-1);C:values("aa"x1,"BB"x2,""x1,"cc"x1)+boundary("aa","BB","BB","","cc");D:int(0,50)+boundary(3,9,5,7,"")
FIXTURE  log_02.xlsx sheet=Use cols=A:Filler Class,B:Total Qty Issued rows=0 stale=1
TEST     fixture=log_01.xlsx call=BuildRecentUsage(Log, Use) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTA(Use![Filler Class]) = 6 AND COUNTIF(Use![Filler Class], "stale") = 0 AND SUM(Use![Total Qty Issued]) = 78
TEST     fixture=log_01.xlsx call=BuildRecentUsage(Log, Use) assert=VALUE(Use![Filler Class]2) = "STICK" AND VALUE(Use![Total Qty Issued]2) = 28 AND VALUE(Use![Filler Class]3) = "TIG ROD" AND VALUE(Use![Total Qty Issued]3) = 26 AND VALUE(Use![Filler Class]4) = "MIG WIRE" AND VALUE(Use![Total Qty Issued]4) = 16 AND VALUE(Use![Filler Class]5) = "FLUX" AND VALUE(Use![Total Qty Issued]5) = 8 AND VALUE(Use![Filler Class]6) = "GAS" AND VALUE(Use![Total Qty Issued]6) = 0 AND COUNTA(Log![Issue No]) = PRE_COUNTA(Log![Issue No]) AND SUM(Log![Qty Issued]) = PRE_SUM(Log![Qty Issued])
TEST     fixture=log_02.xlsx call=BuildRecentUsage(Log, Use) assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(Use![Filler Class]2) = "BB" AND VALUE(Use![Total Qty Issued]2) = 9 AND VALUE(Use![Filler Class]3) = "AA" AND VALUE(Use![Total Qty Issued]3) = 3
TEST     fixture=log_01.xlsx call=BuildRecentUsage(Use, Log) assert=RESULT = -1 AND ERROR_COUNT = 1
