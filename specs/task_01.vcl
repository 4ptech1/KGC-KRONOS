VCL 0.2
TASK
# ASSUME: BOM headers (row 1) include "Part No" and "Required Qty"; GRN headers include "Part No" and "Received Qty"; column order does not matter
# ASSUME: a part can appear on several GRN rows; its received quantities are added
# ASSUME: the Shortages sheet already exists; it is cleared and rebuilt on every run
# ASSUME: a blank or non-numeric Received Qty is ignored; a part with no valid GRN rows counts as 0 received
# ASSUME: a blank or non-numeric Required Qty is listed as "Required qty missing" instead of being skipped
PROC     BuildShortageList
KIND     Function
MODULE   modMaterials
ARGS     bom As Worksheet, grn As Worksheet, out As Worksheet
RET      Long
VAR      received As Dictionary
VAR      bomPart As Long
VAR      bomReq As Long
VAR      grnPart As Long
VAR      grnQty As Long
VAR      partNo As String
VAR      got As Double
VAR      outRow As Long
PRE      COL($bom, "Part No") > 0; COL($bom, "Required Qty") > 0; COL($grn, "Part No") > 0; COL($grn, "Received Qty") > 0
MODIFIES $out
OP       LET(bomPart, COL($bom, "Part No"))
OP       LET(bomReq, COL($bom, "Required Qty"))
OP       LET(grnPart, COL($grn, "Part No"))
OP       LET(grnQty, COL($grn, "Received Qty"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Part No", "Required", "Received", "Shortage", "Remark")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($grn, grnPart) var=r
OP         IF NOT IS_EMPTY(CELL($grn, r, grnPart)) AND IS_NUMBER(CELL($grn, r, grnQty)) THEN DICT_ADD(received, TRIM(CELL($grn, r, grnPart)), CELL($grn, r, grnQty), mode=sum)
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($bom, bomPart) var=r
OP         SKIP_IF IS_EMPTY(CELL($bom, r, bomPart))
OP         LET(partNo, TRIM(CELL($bom, r, bomPart)))
OP         LET(got, 0)
OP         IF EXISTS(received, partNo) THEN DICT_GET(received, partNo, into=got)
OP         IF NOT IS_NUMBER(CELL($bom, r, bomReq))
OP           SET_VALUE(CELL($out, outRow, 1), partNo)
OP           SET_VALUE(CELL($out, outRow, 3), got)
OP           SET_VALUE(CELL($out, outRow, 5), "Required qty missing")
OP           LET(outRow, outRow + 1)
OP         ELSE
OP           IF got < CELL($bom, r, bomReq)
OP             WRITE_ROW($out, outRow, partNo, CELL($bom, r, bomReq), got, CELL($bom, r, bomReq) - got)
OP             IF got = 0 THEN SET_VALUE(CELL($out, outRow, 5), "Not received")
OP             IF got > 0 THEN SET_VALUE(CELL($out, outRow, 5), "Short")
OP             LET(outRow, outRow + 1)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  bom_grn_01.xlsx sheet=BOM cols=A:Part No,B:Description,C:Required Qty rows=40 data=A:seq(1001,1)+blank(2);B:text(12);C:int(1,50)+boundary(30,7)+blank(3)
FIXTURE  bom_grn_01.xlsx sheet=GRN cols=A:Part No,B:Received Qty rows=35 data=A:seq(1001,1)+boundary(1001,1002)+dupes(4);B:int(1,50)+boundary(1,7)+blank(4)
FIXTURE  bom_grn_01.xlsx sheet=Shortages cols=A:Part No,B:Required,C:Received,D:Shortage,E:Remark rows=0 stale=40
TEST     fixture=bom_grn_01.xlsx call=BuildShortageList(BOM, GRN, Shortages) runs=2 assert=RESULT >= 5 AND COUNTA(Shortages![Part No]) = RESULT + 1 AND COUNTIF(Shortages![Part No], "stale") = 0 AND COUNTIF(Shortages![Shortage], "<0") = 0 AND ERROR_COUNT = 0
TEST     fixture=bom_grn_01.xlsx call=BuildShortageList(BOM, GRN, Shortages) assert=COUNTIF(Shortages![Remark], "Required qty missing") = 3 AND COUNTIFS(Shortages![Part No], 1001, Shortages![Remark], "Short") = 1 AND COUNTIF(Shortages![Part No], 1002) = 0
TEST     fixture=bom_grn_01.xlsx call=BuildShortageList(BOM, GRN, Shortages) assert=SUM(BOM![Required Qty]) = PRE_SUM(BOM![Required Qty]) AND SUM(GRN![Received Qty]) = PRE_SUM(GRN![Received Qty])
TEST     fixture=bom_grn_01.xlsx call=BuildShortageList(Shortages, GRN, Shortages) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Shortages![Part No]) = PRE_COUNTA(Shortages![Part No])
