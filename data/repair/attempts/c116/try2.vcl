VCL 0.6
TASK
# ASSUME: the sheets are "Calibration Tracker" and "Location Due"; Location Due exists and is rebuilt on every run
# ASSUME: headers in row 1 of the tracker include "Instrument ID", "Location" and "Next Calibration Due"
# ASSUME: Locations are matched after trimming and ignoring case; the first spelling found is shown
# ASSUME: overdue means a real Next Calibration Due date before today; a blank or non-date due date is not overdue
# ASSUME: rows with a blank Location are ignored, and rows are counted by Instrument ID rows
GOAL     Rebuild the Location Due sheet with each location, how many instruments it has and how many are overdue for calibration, locations A to Z.
PROC     BuildLocationDue
KIND     Function
MODULE   modCalibration
ARGS     cal As Worksheet, out As Worksheet
RET      Long
VAR      cId As Long
VAR      cLoc As Long
VAR      cDue As Long
VAR      total As Dictionary
VAR      late As Dictionary
VAR      nInst As Long
VAR      nLate As Long
VAR      outRow As Long
PRE      COL($cal, "Instrument ID") > 0; COL($cal, "Location") > 0; COL($cal, "Next Calibration Due") > 0
MODIFIES $out
OP       LET(cId, COL($cal, "Instrument ID"))
OP       LET(cLoc, COL($cal, "Location"))
OP       LET(cDue, COL($cal, "Next Calibration Due"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Location", "Instruments", "Overdue")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($cal, cId) var=r
OP         SKIP_IF IS_EMPTY(CELL($cal, r, cLoc))
OP         DICT_ADD(total, TRIM(CELL($cal, r, cLoc)), 1, mode=sum)
OP         DICT_ADD(late, TRIM(CELL($cal, r, cLoc)), 0, mode=sum)
OP         IF IS_DATE(CELL($cal, r, cDue))
OP           IF CELL($cal, r, cDue) < TODAY
OP             DICT_ADD(late, TRIM(CELL($cal, r, cLoc)), 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(total)
OP         DICT_GET(total, k, into=nInst)
OP         DICT_GET(late, k, into=nLate)
OP         WRITE_ROW($out, outRow, k, nInst, nLate)
OP         LET(outRow, outRow + 1)
OP       IF outRow > 3
OP         SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  cal_01.xlsx sheet=Cal cols=A:Instrument ID,B:Location,C:Next Calibration Due rows=14 data=A:seq(1,1);B:values("Lab"x3,"lab "x1,"Bay 2"x3,"Yard"x2,"Store"x2,""x2,"Aaa"x1)+boundary("Lab","Lab","lab ","Bay 2","Bay 2","Bay 2","Yard","Yard","Store","","Store","Aaa","","Lab");C:reldates(-60,60)+boundary(-1,0,5,-30,"",-1,1,10,-5,-9,"",3,"",2)
FIXTURE  cal_01.xlsx sheet=Due cols=A:Location,B:Instruments,C:Overdue rows=0 stale=3
FIXTURE  cal_02.xlsx sheet=Cal cols=A:Instrument ID,B:Location,C:Next Calibration Due rows=3 data=A:seq(1,1);B:values("Zed"x2,"Alpha"x1)+boundary("Zed","Alpha","Zed");C:reldates(-5,5)+boundary(-1,3,2)
FIXTURE  cal_02.xlsx sheet=Due cols=A:Location,B:Instruments,C:Overdue rows=0 stale=1
TEST     fixture=cal_01.xlsx call=BuildLocationDue(Cal, Due) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTA(Due![Location]) = 6 AND COUNTIF(Due![Location], "stale") = 0 AND SUM(Due![Instruments]) = 12 AND SUM(Due![Overdue]) = 4
TEST     fixture=cal_01.xlsx call=BuildLocationDue(Cal, Due) assert=VALUE(Due![Location]2) = "Aaa" AND VALUE(Due![Instruments]2) = 1 AND VALUE(Due![Overdue]2) = 0 AND VALUE(Due![Location]3) = "Bay 2" AND VALUE(Due![Instruments]3) = 3 AND VALUE(Due![Overdue]3) = 2 AND VALUE(Due![Location]4) = "Lab" AND VALUE(Due![Instruments]4) = 4 AND VALUE(Due![Overdue]4) = 1 AND VALUE(Due![Location]5) = "Store" AND VALUE(Due![Instruments]5) = 2 AND VALUE(Due![Overdue]5) = 1 AND VALUE(Due![Location]6) = "Yard" AND VALUE(Due![Instruments]6) = 2 AND VALUE(Due![Overdue]6) = 0 AND COUNTA(Cal![Instrument ID]) = PRE_COUNTA(Cal![Instrument ID]) AND COUNTA(Cal![Location]) = PRE_COUNTA(Cal![Location])
TEST     fixture=cal_02.xlsx call=BuildLocationDue(Cal, Due) assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(Due![Location]2) = "Alpha" AND VALUE(Due![Instruments]2) = 1 AND VALUE(Due![Overdue]2) = 0 AND VALUE(Due![Location]3) = "Zed" AND VALUE(Due![Instruments]3) = 2 AND VALUE(Due![Overdue]3) = 1
TEST     fixture=cal_01.xlsx call=BuildLocationDue(Due, Cal) assert=RESULT = -1 AND ERROR_COUNT = 1
