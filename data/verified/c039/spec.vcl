VCL 0.6
TASK
# ASSUME: BOM headers in row 1 include "Part No" and "Required Qty"; GRN headers include "Part No" and "Received Qty"
# ASSUME: the BOM sheet is named BOM and the GRN sheet GRN; the new BOM headings are "Received" and "Balance"
# ASSUME: a missing Received or Balance heading is added after the last used column of row 1; an existing one is reused
# ASSUME: Part No is trimmed and matched ignoring case; GRN lines with a blank Part No or a Received Qty that is not a number are ignored
# ASSUME: Received is 0 for a part with no GRN line; Balance is Received minus Required Qty, and is left blank when Required Qty is not a number
# ASSUME: a BOM row with a blank Part No gets blank Received and Balance cells; both columns are recalculated from the GRN on every run
# ASSUME: the last row of the BOM and of the GRN has a Part No
GOAL     Add Received and Balance columns to the BOM with the total received per part from the GRN lines and received minus required.
PROC     AddReceivedAndBalanceToBom
KIND     Function
MODULE   modGrnCheck
ARGS     bom As Worksheet, grn As Worksheet
RET      Long
VAR      cPart As Long
VAR      cReq As Long
VAR      cRec As Long
VAR      cBal As Long
VAR      cGPart As Long
VAR      cGQty As Long
VAR      recd As Dictionary
VAR      got As Long
VAR      done As Long
PRE      COL($bom, "Part No") > 0; COL($bom, "Required Qty") > 0; COL($grn, "Part No") > 0; COL($grn, "Received Qty") > 0
MODIFIES $bom
OP       LET(cPart, COL($bom, "Part No"))
OP       LET(cReq, COL($bom, "Required Qty"))
OP       LET(cGPart, COL($grn, "Part No"))
OP       LET(cGQty, COL($grn, "Received Qty"))
OP       LET(cRec, COL($bom, "Received"))
OP       IF cRec = 0
OP         LET(cRec, LASTCOL($bom, 1) + 1)
OP         SET_VALUE(CELL($bom, 1, cRec), "Received")
OP       LET(cBal, COL($bom, "Balance"))
OP       IF cBal = 0
OP         LET(cBal, LASTCOL($bom, 1) + 1)
OP         SET_VALUE(CELL($bom, 1, cBal), "Balance")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($grn, cGPart) var=g
OP         SKIP_IF IS_EMPTY(CELL($grn, g, cGPart))
OP         IF IS_NUMBER(CELL($grn, g, cGQty))
OP           DICT_ADD(recd, TRIM(CELL($grn, g, cGPart)), CELL($grn, g, cGQty), mode=sum)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($bom, cPart) var=r
OP         IF IS_EMPTY(CELL($bom, r, cPart))
OP           CLEAR(CELL($bom, r, cRec), what=contents)
OP           CLEAR(CELL($bom, r, cBal), what=contents)
OP         ELSE
OP           LET(got, 0)
OP           IF EXISTS(recd, TRIM(CELL($bom, r, cPart)))
OP             DICT_GET(recd, TRIM(CELL($bom, r, cPart)), into=got)
OP           SET_VALUE(CELL($bom, r, cRec), got)
OP           IF IS_NUMBER(CELL($bom, r, cReq))
OP             SET_VALUE(CELL($bom, r, cBal), got - CELL($bom, r, cReq))
OP           ELSE
OP             CLEAR(CELL($bom, r, cBal), what=contents)
OP           LET(done, done + 1)
OP       RETURN(done)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  bom_01.xlsx sheet=BOM cols=A:Part No,B:Required Qty rows=6 data=A:values("P-100"x1,"P-200"x1,"P-300"x1," p-400 "x1,""x1,"P-500"x1)+boundary("P-100","P-200","P-300"," p-400 ","","P-500");B:values("10"x1,"5"x1,"8"x1,"4"x1,"3"x1,"tbc"x1)+boundary("10","5","8","4","3","tbc")
FIXTURE  bom_01.xlsx sheet=GRN cols=A:Part No,B:Received Qty rows=10 data=A:values("P-100"x2,""x1,"P-200"x3,"P-400"x1," P-400 "x1,"P-500"x1,"P-600"x1)+boundary("P-100","P-100","","P-200","P-400"," P-400 ","P-500","P-600","P-200","P-200");B:values("4"x1,"7"x1,"8"x1,"5"x1,"3"x1,"2"x1,"6"x1,"9"x1,"n/a"x1,""x1)+boundary("4","7","8","5","3","2","6","9","n/a","")
FIXTURE  bom_02.xlsx sheet=BOM cols=A:Part No,B:Balance,C:Required Qty,D:Received rows=4 data=A:values("P-100"x1,""x1,"P-300"x1,"P-500"x1)+boundary("P-100","","P-300","P-500");B:int(90,99);C:values("10"x1,"3"x1,"8"x1,"tbc"x1)+boundary("10","3","8","tbc");D:int(90,99)
FIXTURE  bom_02.xlsx sheet=GRN cols=A:Part No,B:Received Qty rows=10 data=A:values("P-100"x2,""x1,"P-200"x3,"P-400"x1," P-400 "x1,"P-500"x1,"P-600"x1)+boundary("P-100","P-100","","P-200","P-400"," P-400 ","P-500","P-600","P-200","P-200");B:values("4"x1,"7"x1,"8"x1,"5"x1,"3"x1,"2"x1,"6"x1,"9"x1,"n/a"x1,""x1)+boundary("4","7","8","5","3","2","6","9","n/a","")
TEST     fixture=bom_01.xlsx call=AddReceivedAndBalanceToBom(BOM, GRN) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTA(BOM!A1:D1) = 4 AND COUNTIF(BOM!A1:F1, "Received") = 1 AND COUNTIF(BOM!A1:F1, "Balance") = 1 AND VALUE(BOM![Received]2) = 11 AND VALUE(BOM![Balance]2) = 1 AND VALUE(BOM![Received]3) = 5 AND VALUE(BOM![Balance]3) = 0 AND VALUE(BOM![Received]4) = 0 AND VALUE(BOM![Balance]4) = -8 AND VALUE(BOM![Received]5) = 5 AND VALUE(BOM![Balance]5) = 1 AND VALUE(BOM![Received]7) = 6
TEST     fixture=bom_01.xlsx call=AddReceivedAndBalanceToBom(BOM, GRN) assert=COUNTA(BOM![Received]) = 6 AND COUNTA(BOM![Balance]) = 5 AND SUM(BOM![Received]) = 27 AND SUM(BOM![Balance]) = -6 AND COUNTA(BOM![Part No]) = PRE_COUNTA(BOM![Part No]) AND COUNTA(BOM![Required Qty]) = PRE_COUNTA(BOM![Required Qty]) AND COUNTA(GRN![Part No]) = PRE_COUNTA(GRN![Part No]) AND SUM(GRN![Received Qty]) = PRE_SUM(GRN![Received Qty]) AND ERROR_COUNT = 0
TEST     fixture=bom_02.xlsx call=AddReceivedAndBalanceToBom(BOM, GRN) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTA(BOM!A1:D1) = 4 AND COUNTIF(BOM!A1:H1, "Received") = 1 AND COUNTIF(BOM!A1:H1, "Balance") = 1 AND COUNTA(BOM![Received]) = 4 AND COUNTA(BOM![Balance]) = 3 AND VALUE(BOM![Received]2) = 11 AND VALUE(BOM![Balance]2) = 1 AND VALUE(BOM![Received]4) = 0 AND VALUE(BOM![Balance]4) = -8 AND VALUE(BOM![Received]5) = 6 AND SUM(BOM![Received]) = 17 AND SUM(BOM![Balance]) = -7 AND COUNTA(BOM![Part No]) = PRE_COUNTA(BOM![Part No]) AND COUNTA(BOM![Required Qty]) = PRE_COUNTA(BOM![Required Qty])
TEST     fixture=bom_01.xlsx call=AddReceivedAndBalanceToBom(GRN, BOM) assert=RESULT = -1 AND ERROR_COUNT = 1
