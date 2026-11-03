VCL 0.6
TASK
# ASSUME: GRN headers in row 1 include "GRN No", "Part No", "Ordered Qty" and "Received Qty", and the data starts at A1
# ASSUME: the real sheets are "GRN" and "Short Receipts"; Short Receipts already exists and is cleared and rebuilt on every run; the TESTs use GRN and ShortReceipts
# ASSUME: a line is short only when both quantities are numbers and Received Qty is below Ordered Qty; an equal or larger Received Qty is not listed
# ASSUME: lines with a blank or non-numeric Ordered Qty or Received Qty are skipped; a quantity typed as text such as "20" counts as the number 20
# ASSUME: Short Receipts gets the headings GRN No, Part No, Ordered Qty, Received Qty and Shortfall in row 1; Shortfall is Ordered Qty minus Received Qty
# ASSUME: the biggest shortfall comes first; lines with the same shortfall have no fixed order
# ASSUME: the function returns how many short lines were listed
GOAL     List every GRN line received short of what was ordered on the Short Receipts sheet with its shortfall, biggest shortfall first.
PROC     ListShortReceipts
KIND     Function
MODULE   modReceipts
ARGS     grn As Worksheet, out As Worksheet
RET      Long
VAR      cGrn As Long
VAR      cPart As Long
VAR      cOrd As Long
VAR      cRec As Long
VAR      data As Variant
VAR      gap As Double
VAR      outRow As Long
PRE      COL($grn, "GRN No") > 0; COL($grn, "Part No") > 0; COL($grn, "Ordered Qty") > 0; COL($grn, "Received Qty") > 0
MODIFIES $out
OP       LET(cGrn, COL($grn, "GRN No"))
OP       LET(cPart, COL($grn, "Part No"))
OP       LET(cOrd, COL($grn, "Ordered Qty"))
OP       LET(cRec, COL($grn, "Received Qty"))
OP       READ_ARRAY(USED_RANGE($grn), into=data)
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "GRN No", "Part No", "Ordered Qty", "Received Qty", "Shortfall")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=ROWS(data) var=r
OP         IF IS_NUMBER(ARR(data, r, cOrd))
OP           IF IS_NUMBER(ARR(data, r, cRec))
OP             LET(gap, ARR(data, r, cOrd) - ARR(data, r, cRec))
OP             IF gap > 0
OP               WRITE_ROW($out, outRow, ARR(data, r, cGrn), ARR(data, r, cPart), ARR(data, r, cOrd), ARR(data, r, cRec), gap)
OP               LET(outRow, outRow + 1)
OP       IF outRow > 2
OP         SORT(USED_RANGE($out), key=5, order=desc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  grn_01.xlsx sheet=GRN cols=A:GRN No,B:Part No,C:Ordered Qty,D:Received Qty rows=8 data=A:seq(101,1);B:values("BR-10"x1,"BR-20"x1,"FL-30"x1,"FL-40"x1,"GS-50"x1,"GS-60"x1,"NZ-70"x1,"NZ-80"x1)+boundary("BR-10","BR-20","FL-30","FL-40","GS-50","GS-60","NZ-70","NZ-80");C:int(5,50)+boundary(10,10,10,10,20,"",50,8);D:int(0,9)+boundary(4,10,9,12,"",3,30,0)
FIXTURE  grn_01.xlsx sheet=ShortReceipts cols=A:GRN No,B:Part No,C:Ordered Qty,D:Received Qty,E:Shortfall rows=0 stale=3
FIXTURE  grn_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  grn_02.xlsx sheet=GRN cols=A:GRN No,B:Part No,C:Ordered Qty,D:Received Qty rows=7 data=A:seq(201,1);B:text(6);C:values("TBC"x1,"20"x2,"15"x1,"40"x1,"9"x1,"100"x1)+boundary("TBC","20","20","15","40","9","100");D:values("3"x1,"n/a"x1,"12"x1,"15"x1,"40"x1,"10"x1,"20"x1)+boundary("3","n/a","12","15","40","10","20")
FIXTURE  grn_02.xlsx sheet=ShortReceipts cols=A:GRN No,B:Part No,C:Ordered Qty,D:Received Qty,E:Shortfall rows=0 stale=3
FIXTURE  grn_03.xlsx sheet=GRN cols=A:GRN No,B:Part No,C:Ordered Qty,D:Received Qty rows=3 data=A:seq(301,1);B:text(6);C:int(5,50)+boundary(10,25,30);D:int(0,9)+boundary(10,26,30)
FIXTURE  grn_03.xlsx sheet=ShortReceipts cols=A:GRN No,B:Part No,C:Ordered Qty,D:Received Qty,E:Shortfall rows=0 stale=3
TEST     fixture=grn_01.xlsx call=ListShortReceipts(GRN, ShortReceipts) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(ShortReceipts![GRN No], "stale") = 0 AND ROWS_USED(ShortReceipts) = 5 AND COUNTA(ShortReceipts![GRN No]) = 5 AND SUM(ShortReceipts![Shortfall]) = 35
TEST     fixture=grn_01.xlsx call=ListShortReceipts(GRN, ShortReceipts) assert=RESULT = 4 AND VALUE(ShortReceipts![GRN No]2) = 107 AND VALUE(ShortReceipts![Part No]2) = "NZ-70" AND VALUE(ShortReceipts![Ordered Qty]2) = 50 AND VALUE(ShortReceipts![Received Qty]2) = 30 AND VALUE(ShortReceipts![Shortfall]2) = 20 AND VALUE(ShortReceipts![GRN No]3) = 108 AND VALUE(ShortReceipts![Shortfall]3) = 8 AND VALUE(ShortReceipts![GRN No]4) = 101 AND VALUE(ShortReceipts![Shortfall]4) = 6 AND VALUE(ShortReceipts![GRN No]5) = 103 AND VALUE(ShortReceipts![Shortfall]5) = 1
TEST     fixture=grn_01.xlsx call=ListShortReceipts(GRN, ShortReceipts) assert=ERROR_COUNT = 0 AND COUNTA(GRN![GRN No]) = PRE_COUNTA(GRN![GRN No]) AND SUM(GRN![Ordered Qty]) = PRE_SUM(GRN![Ordered Qty]) AND SUM(GRN![Received Qty]) = PRE_SUM(GRN![Received Qty])
TEST     fixture=grn_01.xlsx call=ListShortReceipts(Notes, ShortReceipts) assert=RESULT = -1 AND ERROR_COUNT = 1
TEST     fixture=grn_02.xlsx call=ListShortReceipts(GRN, ShortReceipts) assert=RESULT = 2 AND ERROR_COUNT = 0 AND ROWS_USED(ShortReceipts) = 3 AND VALUE(ShortReceipts![GRN No]2) = 207 AND VALUE(ShortReceipts![Shortfall]2) = 80 AND VALUE(ShortReceipts![GRN No]3) = 203 AND VALUE(ShortReceipts![Shortfall]3) = 8 AND COUNTA(GRN![GRN No]) = PRE_COUNTA(GRN![GRN No])
TEST     fixture=grn_03.xlsx call=ListShortReceipts(GRN, ShortReceipts) assert=RESULT = 0 AND ERROR_COUNT = 0 AND ROWS_USED(ShortReceipts) = 1 AND COUNTA(ShortReceipts![GRN No]) = 1 AND COUNTA(GRN![GRN No]) = PRE_COUNTA(GRN![GRN No])
