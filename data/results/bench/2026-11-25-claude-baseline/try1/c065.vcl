VCL 0.7
TASK
# ASSUME: GRN headers in row 1 include "GRN No", "Part No", "Ordered Qty" and "Received Qty"; Short Receipts is an existing sheet rebuilt on every run
# ASSUME: lines with a blank or non-numeric Ordered Qty or Received Qty are skipped; ties keep no particular order
GOAL     Rebuild the Short Receipts sheet with every GRN line received short, with a Shortfall column, biggest shortfall first.
PROC     ListShortReceipts
KIND     Function
MODULE   modReceipts
ARGS     grn As Worksheet, out As Worksheet
RET      Long
VAR      cNo As Long
VAR      cPart As Long
VAR      cOrd As Long
VAR      cRec As Long
VAR      outRow As Long
VAR      listed As Long
PRE      COL($grn, "GRN No") > 0; COL($grn, "Part No") > 0; COL($grn, "Ordered Qty") > 0; COL($grn, "Received Qty") > 0
MODIFIES $out
OP       LET(cNo, COL($grn, "GRN No"))
OP       LET(cPart, COL($grn, "Part No"))
OP       LET(cOrd, COL($grn, "Ordered Qty"))
OP       LET(cRec, COL($grn, "Received Qty"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "GRN No", "Part No", "Ordered Qty", "Received Qty", "Shortfall")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($grn, cNo) var=r
OP         SKIP_IF NOT IS_NUMBER(CELL($grn, r, cOrd))
OP         SKIP_IF NOT IS_NUMBER(CELL($grn, r, cRec))
OP         IF CELL($grn, r, cRec) < CELL($grn, r, cOrd)
OP           WRITE_ROW($out, outRow, CELL($grn, r, cNo), CELL($grn, r, cPart), CELL($grn, r, cOrd), CELL($grn, r, cRec), CELL($grn, r, cOrd) - CELL($grn, r, cRec))
OP           LET(outRow, outRow + 1)
OP           LET(listed, listed + 1)
OP       SORT(USED_RANGE($out), key=5, order=desc, header=yes)
OP       RETURN(listed)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  grn_01.xlsx sheet=Grn cols=A:GRN No,B:Part No,C:Ordered Qty,D:Received Qty rows=8 data=A:seq(1,1);B:text(6);C:int(1,100)+boundary(10,10,20,40,5,10,30,50);D:values("8"x1,"10"x1,"12"x1,"30"x1,"5"x1,"n/a"x1,""x1,"60"x1)+boundary("8","10","12","30","5","n/a","","60")
FIXTURE  grn_01.xlsx sheet=Short cols=A:GRN No,B:Part No,C:Ordered Qty,D:Received Qty,E:Shortfall rows=0 stale=2
FIXTURE  grn_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=grn_01.xlsx call=ListShortReceipts(Grn, Short) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTA(Short![GRN No]) = 4 AND COUNTIF(Short![GRN No], "stale") = 0 AND VALUE(Short![Shortfall]2) = 10 AND VALUE(Short![Shortfall]3) = 8 AND VALUE(Short![Shortfall]4) = 2
TEST     fixture=grn_01.xlsx call=ListShortReceipts(Grn, Short) assert=COUNTA(Grn![GRN No]) = PRE_COUNTA(Grn![GRN No]) AND COUNTA(Grn![Received Qty]) = PRE_COUNTA(Grn![Received Qty]) AND VALUE(Short![GRN No]2) = 4 AND VALUE(Short![GRN No]3) = 3 AND VALUE(Short![GRN No]4) = 1
TEST     fixture=grn_01.xlsx call=ListShortReceipts(Notes, Short) assert=RESULT = -1 AND ERROR_COUNT = 1
