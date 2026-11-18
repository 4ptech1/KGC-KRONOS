VCL 0.6
TASK
# ASSUME: Weld Log headers in row 1 include "Weld Seam No", "Welder ID" and "Visual"; the Welder Visual sheet exists and is passed in, and is rebuilt on every run
# ASSUME: Welder IDs are trimmed and matched ignoring case; rows with a blank Welder ID are ignored
# ASSUME: a weld counts as not Accepted when its Visual is anything but Accepted (ignoring case and spaces), a blank Visual included
# ASSUME: welders with the same number of welds keep the order in which SORT leaves them; the result is the number of welders written
GOAL     Rebuild the Welder Visual sheet with each welder, how many welds they made and how many were not accepted, the welder with the most welds first.
PROC     BuildWelderVisual
KIND     Function
MODULE   modWelderVisual
ARGS     weld As Worksheet, out As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cWelder As Long
VAR      cVisual As Long
VAR      total As Dictionary
VAR      bad As Dictionary
VAR      nWelds As Long
VAR      nBad As Long
VAR      outRow As Long
PRE      COL($weld, "Weld Seam No") > 0; COL($weld, "Welder ID") > 0; COL($weld, "Visual") > 0
MODIFIES $out
OP       LET(cSeam, COL($weld, "Weld Seam No"))
OP       LET(cWelder, COL($weld, "Welder ID"))
OP       LET(cVisual, COL($weld, "Visual"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Welder ID", "Welds", "Not Accepted")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($weld, cSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($weld, r, cWelder))
OP         DICT_ADD(total, TRIM(CELL($weld, r, cWelder)), 1, mode=sum)
OP         DICT_ADD(bad, TRIM(CELL($weld, r, cWelder)), IIF(SAME(CELL($weld, r, cVisual), "Accepted"), 0, 1), mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(total)
OP         DICT_GET(total, k, into=nWelds)
OP         DICT_GET(bad, k, into=nBad)
OP         WRITE_ROW($out, outRow, k, nWelds, nBad)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=2, order=desc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off
ERR      standard
FIXTURE  welder_01.xlsx sheet=Weld cols=A:Weld Seam No,B:Welder ID,C:Visual rows=12 data=A:seq(1,1);B:values(" W-3 "x1,"W-1"x2,""x1,"W-2"x2,"w-1"x1,"W-4"x1," W-1"x1,"W-2 "x1,"W-3"x1,"W-1 "x1)+boundary(" W-3 ","W-1","","W-2","w-1","W-4"," W-1","W-2 ","W-3","W-1 ","W-2","W-1");C:values("Accepted"x7,"Rejected"x2,"accepted"x1,"Repair"x1,""x1)+boundary("Accepted","Rejected","Accepted","accepted","Accepted","Repair","Accepted","","Accepted","Rejected","Accepted","Accepted")
FIXTURE  welder_01.xlsx sheet=Out cols=A:Welder ID,B:Welds,C:Not Accepted rows=0 stale=3
TEST     fixture=welder_01.xlsx call=BuildWelderVisual(Weld, Out) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(Out![Welder ID]) = 5 AND COUNTIF(Out![Welder ID], "stale") = 0 AND VALUE(Out![Welder ID]2) = "W-1" AND VALUE(Out![Welds]2) = 5 AND VALUE(Out![Not Accepted]2) = 2 AND VALUE(Out![Welder ID]3) = "W-2" AND VALUE(Out![Welds]3) = 3 AND VALUE(Out![Not Accepted]3) = 1
TEST     fixture=welder_01.xlsx call=BuildWelderVisual(Weld, Out) assert=COUNTA(Weld![Weld Seam No]) = PRE_COUNTA(Weld![Weld Seam No]) AND COUNTA(Weld![Welder ID]) = PRE_COUNTA(Weld![Welder ID]) AND VALUE(Out![Welder ID]4) = "W-3" AND VALUE(Out![Welds]4) = 2 AND VALUE(Out![Not Accepted]4) = 0 AND VALUE(Out![Welder ID]5) = "W-4" AND VALUE(Out![Welds]5) = 1 AND VALUE(Out![Not Accepted]5) = 1 AND ROWS_USED(Out) = 5
TEST     fixture=welder_01.xlsx call=BuildWelderVisual(Out, Weld) assert=RESULT = -1 AND ERROR_COUNT = 1
