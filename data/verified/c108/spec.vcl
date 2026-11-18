VCL 0.6
TASK
# ASSUME: GRN headers in row 1 are GRN No, Part No, Received Qty; BOM headers are Part No, Required Qty, Received Total, Remark
# ASSUME: part numbers are matched trimmed and ignoring case; BOM rows with a blank Part No are left untouched
# ASSUME: a GRN Received Qty that is not a number (blank or text) adds nothing; GRN parts not in the BOM are ignored
# ASSUME: a part is Short only when its Required Qty is a number and the received total is less than it; any other part gets its Remark cleared
GOAL     Total the received quantity of each BOM part from the GRN sheet, mark the parts received short of the required quantity, and return how many are Short.
PROC     MarkShortParts
KIND     Function
MODULE   modReceiving
ARGS     grn As Worksheet, bom As Worksheet
RET      Long
VAR      gPart As Long
VAR      gQty As Long
VAR      cPart As Long
VAR      cReq As Long
VAR      cRec As Long
VAR      cRemark As Long
VAR      rec As Dictionary
VAR      q As Double
VAR      got As Double
VAR      need As Double
VAR      shortN As Long
PRE      COL($grn, "Part No") > 0; COL($grn, "Received Qty") > 0; COL($bom, "Part No") > 0; COL($bom, "Required Qty") > 0; COL($bom, "Received Total") > 0; COL($bom, "Remark") > 0
MODIFIES $bom["Received Total", "Remark"]
OP       LET(gPart, COL($grn, "Part No"))
OP       LET(gQty, COL($grn, "Received Qty"))
OP       LET(cPart, COL($bom, "Part No"))
OP       LET(cReq, COL($bom, "Required Qty"))
OP       LET(cRec, COL($bom, "Received Total"))
OP       LET(cRemark, COL($bom, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($grn, gPart) var=g
OP         IF IS_NUMBER(CELL($grn, g, gQty))
OP           LET(q, CELL($grn, g, gQty))
OP           DICT_ADD(rec, TRIM(CELL($grn, g, gPart)), q, mode=sum)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($bom, cPart) var=b
OP         SKIP_IF IS_EMPTY(CELL($bom, b, cPart))
OP         LET(got, 0)
OP         DICT_GET(rec, TRIM(CELL($bom, b, cPart)), into=got)
OP         SET_VALUE(CELL($bom, b, cRec), got)
OP         IF IS_NUMBER(CELL($bom, b, cReq))
OP           LET(need, CELL($bom, b, cReq))
OP           IF got < need
OP             SET_VALUE(CELL($bom, b, cRemark), "Short")
OP             LET(shortN, shortN + 1)
OP           ELSE
OP             CLEAR(CELL($bom, b, cRemark), what=contents)
OP         ELSE
OP           CLEAR(CELL($bom, b, cRemark), what=contents)
OP       MSG("Parts short: " & shortN)
OP       RETURN(shortN)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  recv_01.xlsx sheet=GRN cols=A:GRN No,B:Part No,C:Received Qty rows=11 data=A:seq(1,1);B:text(8)+boundary("p-100"," P-100 ","P-200","P-200","P-400 ","","P-600","P-700","p-700","P-999","P-300");C:text(3)+boundary("4","6","5","n/a","7","9","3","2","0.5","8","")
FIXTURE  recv_01.xlsx sheet=BOM cols=A:Part No,B:Required Qty,C:Received Total,D:Remark rows=8 data=A:text(8)+boundary("P-100","P-200","P-300","p-400","P-500","","P-600","P-700");B:text(3)+boundary("10","8","5","5","0","5","n/a","3");C:values("old"x8);D:values("old"x8)
TEST     fixture=recv_01.xlsx call=MarkShortParts(GRN, BOM) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(BOM![Remark], "Short") = 3 AND COUNTIF(BOM![Remark], "old") = 1 AND COUNTA(BOM![Remark]) = 5 AND COUNTIF(BOM![Received Total], "old") = 1 AND SUM(BOM![Received Total]) = 27.5
TEST     fixture=recv_01.xlsx call=MarkShortParts(GRN, BOM) assert=VALUE(BOM![Received Total]2) = 10 AND VALUE(BOM![Received Total]3) = 5 AND VALUE(BOM![Received Total]4) = 0 AND VALUE(BOM![Received Total]5) = 7 AND VALUE(BOM![Received Total]6) = 0 AND VALUE(BOM![Received Total]8) = 3 AND VALUE(BOM![Received Total]9) = 2.5 AND VALUE(BOM![Remark]3) = "Short" AND VALUE(BOM![Remark]4) = "Short" AND VALUE(BOM![Remark]9) = "Short"
TEST     fixture=recv_01.xlsx call=MarkShortParts(GRN, BOM) assert=COUNTA(BOM![Part No]) = PRE_COUNTA(BOM![Part No]) AND COUNTA(GRN![Part No]) = PRE_COUNTA(GRN![Part No]) AND SUM(BOM![Required Qty]) = PRE_SUM(BOM![Required Qty]) AND COUNTA(GRN![Received Qty]) = PRE_COUNTA(GRN![Received Qty])
TEST     fixture=recv_01.xlsx call=MarkShortParts(BOM, GRN) assert=RESULT = -1 AND ERROR_COUNT = 1
