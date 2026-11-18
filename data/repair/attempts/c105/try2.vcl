VCL 0.6
TASK
# ASSUME: GRN sheet headers in row 1 include "GRN No", "Supplier" and "Received Qty"
# ASSUME: each row with a supplier is one GRN line; suppliers are trimmed and matched ignoring case, shown as first spelled
# ASSUME: a Received Qty that is not a number adds nothing to the total, but its line is still counted
# ASSUME: the Supplier Receipts sheet exists, is passed in and is rebuilt on every run, biggest total first
GOAL     Rebuild the Supplier Receipts sheet with each supplier, its number of GRN lines and its total received quantity, biggest total first.
PROC     SummariseSupplierReceipts
KIND     Function
MODULE   modGrn
ARGS     grn As Worksheet, out As Worksheet
RET      Long
VAR      cSupplier As Long
VAR      cQty As Long
VAR      lines As Dictionary
VAR      totals As Dictionary
VAR      nLines As Long
VAR      total As Double
VAR      outRow As Long
PRE      COL($grn, "Supplier") > 0; COL($grn, "Received Qty") > 0
MODIFIES $out
OP       LET(cSupplier, COL($grn, "Supplier"))
OP       LET(cQty, COL($grn, "Received Qty"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Supplier", "GRN Lines", "Total Received Qty")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($grn, cSupplier) var=r
OP         SKIP_IF IS_EMPTY(CELL($grn, r, cSupplier))
OP         DICT_ADD(lines, TRIM(CELL($grn, r, cSupplier)), 1, mode=sum)
OP         DICT_ADD(totals, TRIM(CELL($grn, r, cSupplier)), 0, mode=sum)
OP         IF IS_NUMBER(CELL($grn, r, cQty))
OP           DICT_ADD(totals, TRIM(CELL($grn, r, cSupplier)), CELL($grn, r, cQty), mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(lines)
OP         DICT_GET(lines, k, into=nLines)
OP         DICT_GET(totals, k, into=total)
OP         WRITE_ROW($out, outRow, k, nLines, total)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=3, order=desc, header=yes)
OP       RETURN(COUNT(lines))
PERF     screen_off
ERR      standard
FIXTURE  grn_01.xlsx sheet=Grn cols=A:GRN No,B:Supplier,C:Received Qty rows=12 data=A:seq(1,1);B:values("Acme Steel"x2,"ACME STEEL "x1,"acme steel"x1,"Boron Ltd"x3,"Cobalt Co"x3,""x2)+boundary("Acme Steel","Boron Ltd","ACME STEEL ","","Cobalt Co","Boron Ltd","acme steel","Cobalt Co","","Boron Ltd","Acme Steel","Cobalt Co");C:text(3)+boundary(100,400,60,500,25,"n/a",15,"",7,30,5,10)
FIXTURE  grn_01.xlsx sheet=Receipts cols=A:Supplier,B:GRN Lines,C:Total Received Qty rows=0 stale=2
TEST     fixture=grn_01.xlsx call=SummariseSupplierReceipts(Grn, Receipts) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTA(Receipts![Supplier]) = 4 AND COUNTIF(Receipts![Supplier], "stale") = 0 AND VALUE(Receipts![Supplier]2) = "Boron Ltd" AND VALUE(Receipts![GRN Lines]2) = 3 AND VALUE(Receipts![Total Received Qty]2) = 430 AND VALUE(Receipts![Supplier]3) = "Acme Steel" AND VALUE(Receipts![GRN Lines]3) = 4 AND VALUE(Receipts![Total Received Qty]3) = 180 AND VALUE(Receipts![Supplier]4) = "Cobalt Co" AND VALUE(Receipts![GRN Lines]4) = 3 AND VALUE(Receipts![Total Received Qty]4) = 35
TEST     fixture=grn_01.xlsx call=SummariseSupplierReceipts(Grn, Receipts) assert=SUM(Receipts![GRN Lines]) = 10 AND SUM(Receipts![Total Received Qty]) = 645 AND COUNTA(Grn![Supplier]) = PRE_COUNTA(Grn![Supplier]) AND SUM(Grn![Received Qty]) = PRE_SUM(Grn![Received Qty])
TEST     fixture=grn_01.xlsx call=SummariseSupplierReceipts(Receipts, Grn) assert=RESULT = -1 AND ERROR_COUNT = 1
