VCL 0.6
BUILD
# ASSUME: GRN headers include "Part No", "Heat No", "Received Qty" and "Remark"; Material Certs has "Heat No"; BOM has "Part No", "Required Qty", "Received" and "Status"
# ASSUME: the real sheets are "GRN", "Material Certs", "BOM" and "Short Parts"; the TESTs use GRN, Certs, BOM and Short
# ASSUME: part and heat numbers are matched ignoring case and extra spaces; a GRN line with a blank Part No is skipped everywhere
# ASSUME: a blank Heat No has no cert row; when a line is both without cert and with a bad quantity, No cert wins
# ASSUME: only positive numeric Received Qty values are added up; a BOM line whose Required Qty is not a number gets Status Check
# ASSUME: each part is on one BOM line, so the Short count is the number of Short BOM lines; Short Parts is rebuilt on every run and exists
GOAL     Flag GRN lines without a cert or with a bad quantity, total received per BOM part, list the Short parts A to Z and count them.
NAME     GrnReceivingCheck
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunGrnRemarks()
SIG      modMain.RunBomTotals()
SIG      modMain.RunShortList()
SIG      modMain.RunHeadless(grn As Worksheet, certs As Worksheet, bom As Worksheet, out As Worksheet) As Long
SIG      modSteps.MarkGrnRemarks(grn As Worksheet, certs As Worksheet) As Long
SIG      modSteps.TotalReceived(bom As Worksheet, grn As Worksheet) As Long
SIG      modSteps.ListShortParts(bom As Worksheet, out As Worksheet) As Long
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.MarkGrnRemarks
FLOW     modMain.RunHeadless -> modSteps.TotalReceived
FLOW     modMain.RunHeadless -> modSteps.ListShortParts
FLOW     modMain.RunGrnRemarks -> modSteps.MarkGrnRemarks
FLOW     modMain.RunBomTotals -> modSteps.TotalReceived
FLOW     modMain.RunShortList -> modSteps.ListShortParts
FIXTURE  grn_01.xlsx sheet=GRN cols=A:GRN No,B:Part No,C:Heat No,D:Received Qty,E:Remark rows=10 data=A:seq(1,1);B:values("P-100"x3," p-100 "x1,"P-200"x1,"P-300"x2,""x1,"p-200"x1,"P-900"x1)+boundary("P-100"," p-100 ","P-200","P-300","","p-200","P-100","P-900","P-300","P-100");C:values("H1"x2," h1 "x1,"H2"x2,"H9"x2,""x1,"H3"x1,"h2"x1)+boundary("H1"," h1 ","H2","H9","H9","","H3","H1","H2","h2");D:values("10"x1,"5"x1,"0"x1,"8"x1,"3"x1,"abc"x1,"-2"x1,""x1,"4"x1,"1"x1)+boundary("10","5","0","8","3","abc","-2","","4","1");E:values("old"x10)
FIXTURE  grn_01.xlsx sheet=Certs cols=A:Heat No,B:MTC No rows=6 data=A:values("H1"x1,"H2"x1,""x1,"H3"x1,"h1"x1,"H2"x1)+boundary("H1","H2","","H3","h1","H2");B:text(6)
FIXTURE  grn_01.xlsx sheet=BOM cols=A:Part No,B:Required Qty,C:Received,D:Status rows=7 data=A:values("P-400"x1,"P-100"x1,""x1," P-200 "x1,"P-300"x1,"P-500"x1,"P-600"x1)+boundary("P-400","P-100",""," P-200 ","P-300","P-500","P-600");B:values("5"x1,"16"x1,"3"x1,"1"x1,"13"x1,"tbd"x1,"0"x1)+boundary("5","16","3","1","13","tbd","0");C:values("old"x7);D:values("old"x7)
FIXTURE  grn_01.xlsx sheet=Short cols=A:Part No,B:Required Qty,C:Received,D:Status rows=0 stale=2
TEST     fixture=grn_01.xlsx call=modMain.RunHeadless(GRN, Certs, BOM, Short) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(GRN![Remark], "No cert") = 2 AND COUNTIF(GRN![Remark], "Check") = 3 AND COUNTIF(GRN![Remark], "old") = 1 AND COUNTA(GRN![Remark]) = 7 AND COUNTIF(BOM![Status], "Short") = 3 AND COUNTIF(BOM![Status], "Complete") = 2 AND COUNTIF(BOM![Status], "Check") = 1 AND COUNTIF(BOM![Status], "old") = 1 AND COUNTA(Short![Part No]) = 4 AND COUNTIF(Short![Part No], "stale") = 0
TEST     fixture=grn_01.xlsx call=modMain.RunHeadless(GRN, Certs, BOM, Short) assert=RESULT = 3 AND COUNTA(GRN![GRN No]) = PRE_COUNTA(GRN![GRN No]) AND COUNTA(BOM![Part No]) = PRE_COUNTA(BOM![Part No]) AND COUNTA(Certs![Heat No]) = PRE_COUNTA(Certs![Heat No]) AND SUM(BOM![Received]) = 28 AND VALUE(BOM![Received]3) = 16 AND VALUE(BOM![Received]6) = 12 AND COUNTIF(BOM![Received], "old") = 1 AND VALUE(Short![Part No]2) = "P-200" AND VALUE(Short![Part No]3) = "P-300" AND VALUE(Short![Part No]4) = "P-400" AND VALUE(Short![Received]3) = 12 AND VALUE(Short![Required Qty]4) = 5
TEST     fixture=grn_01.xlsx call=modMain.RunHeadless(Certs, GRN, BOM, Short) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
TASK
GOAL     Start the GRN receiving check on the workbook's own sheets and tell the user how many parts are Short.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "GRN"), SHEET(THIS_BOOK, "Material Certs"), SHEET(THIS_BOOK, "BOM"), SHEET(THIS_BOOK, "Short Parts"), into=n)
OP       MSG("Parts Short: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run only step 1 on the workbook's own sheets and tell the user how many GRN lines got a remark.
PROC     RunGrnRemarks
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modSteps.MarkGrnRemarks, SHEET(THIS_BOOK, "GRN"), SHEET(THIS_BOOK, "Material Certs"), into=n)
OP       MSG("GRN lines with a remark: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run only step 2 on the workbook's own sheets and tell the user how many BOM lines were totalled.
PROC     RunBomTotals
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modSteps.TotalReceived, SHEET(THIS_BOOK, "BOM"), SHEET(THIS_BOOK, "GRN"), into=n)
OP       MSG("BOM lines totalled: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run only step 3 on the workbook's own sheets and tell the user how many parts are Short.
PROC     RunShortList
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modSteps.ListShortParts, SHEET(THIS_BOOK, "BOM"), SHEET(THIS_BOOK, "Short Parts"), into=n)
OP       MSG("Parts Short: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run all three steps on the four sheets it is given and return how many parts are Short.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     grn As Worksheet, certs As Worksheet, bom As Worksheet, out As Worksheet
RET      Long
VAR      marked As Long
VAR      totalled As Long
VAR      short As Long
PRE      COL($grn, "Part No") > 0; COL($grn, "Heat No") > 0; COL($grn, "Received Qty") > 0; COL($grn, "Remark") > 0; COL($certs, "Heat No") > 0; COL($bom, "Part No") > 0; COL($bom, "Required Qty") > 0; COL($bom, "Received") > 0; COL($bom, "Status") > 0
MODIFIES $grn["Remark"]
MODIFIES $bom["Received", "Status"]
MODIFIES $out
OP       CALL(modSteps.MarkGrnRemarks, $grn, $certs, into=marked)
OP       CALL(modSteps.TotalReceived, $bom, $grn, into=totalled)
OP       CALL(modSteps.ListShortParts, $bom, $out, into=short)
OP       RETURN(short)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write No cert or Check in the GRN Remark column, or clear it, and return how many lines got a remark.
PROC     MarkGrnRemarks
KIND     Function
MODULE   modSteps
ARGS     grn As Worksheet, certs As Worksheet
RET      Long
VAR      cPart As Long
VAR      cHeat As Long
VAR      cQty As Long
VAR      cRemark As Long
VAR      cCert As Long
VAR      known As Dictionary
VAR      flag As String
VAR      n As Long
PRE      COL($grn, "Part No") > 0; COL($grn, "Heat No") > 0; COL($grn, "Received Qty") > 0; COL($grn, "Remark") > 0; COL($certs, "Heat No") > 0
MODIFIES $grn["Remark"]
OP       LET(cPart, COL($grn, "Part No"))
OP       LET(cHeat, COL($grn, "Heat No"))
OP       LET(cQty, COL($grn, "Received Qty"))
OP       LET(cRemark, COL($grn, "Remark"))
OP       LET(cCert, COL($certs, "Heat No"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($certs, cCert) var=r
OP         SKIP_IF IS_EMPTY(CELL($certs, r, cCert))
OP         DICT_ADD(known, CLEAN_TEXT(CELL($certs, r, cCert)), 1, mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($grn, cPart) var=r
OP         SKIP_IF IS_EMPTY(CELL($grn, r, cPart))
OP         LET(flag, "")
OP         IF IS_NUMBER(CELL($grn, r, cQty))
OP           IF CELL($grn, r, cQty) <= 0
OP             LET(flag, "Check")
OP         ELSE
OP           LET(flag, "Check")
OP         IF NOT EXISTS(known, CLEAN_TEXT(CELL($grn, r, cHeat)))
OP           LET(flag, "No cert")
OP         IF flag = ""
OP           CLEAR(CELL($grn, r, cRemark), what=contents)
OP         ELSE
OP           SET_VALUE(CELL($grn, r, cRemark), flag)
OP           LET(n, n + 1)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write in each BOM line the total positive quantity received for its part in the GRN and set Status to Complete, Short or Check.
PROC     TotalReceived
KIND     Function
MODULE   modSteps
ARGS     bom As Worksheet, grn As Worksheet
RET      Long
VAR      cGPart As Long
VAR      cGQty As Long
VAR      cPart As Long
VAR      cReq As Long
VAR      cRec As Long
VAR      cStatus As Long
VAR      sums As Dictionary
VAR      got As Long
VAR      n As Long
PRE      COL($grn, "Part No") > 0; COL($grn, "Received Qty") > 0; COL($bom, "Part No") > 0; COL($bom, "Required Qty") > 0; COL($bom, "Received") > 0; COL($bom, "Status") > 0
MODIFIES $bom["Received", "Status"]
OP       LET(cGPart, COL($grn, "Part No"))
OP       LET(cGQty, COL($grn, "Received Qty"))
OP       LET(cPart, COL($bom, "Part No"))
OP       LET(cReq, COL($bom, "Required Qty"))
OP       LET(cRec, COL($bom, "Received"))
OP       LET(cStatus, COL($bom, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($grn, cGPart) var=r
OP         SKIP_IF IS_EMPTY(CELL($grn, r, cGPart))
OP         IF IS_NUMBER(CELL($grn, r, cGQty))
OP           IF CELL($grn, r, cGQty) > 0
OP             DICT_ADD(sums, CLEAN_TEXT(CELL($grn, r, cGPart)), CELL($grn, r, cGQty), mode=sum)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($bom, cPart) var=r
OP         SKIP_IF IS_EMPTY(CELL($bom, r, cPart))
OP         LET(got, 0)
OP         IF EXISTS(sums, CLEAN_TEXT(CELL($bom, r, cPart)))
OP           DICT_GET(sums, CLEAN_TEXT(CELL($bom, r, cPart)), into=got)
OP         SET_VALUE(CELL($bom, r, cRec), got)
OP         IF IS_NUMBER(CELL($bom, r, cReq))
OP           IF got >= CELL($bom, r, cReq)
OP             SET_VALUE(CELL($bom, r, cStatus), "Complete")
OP           ELSE
OP             SET_VALUE(CELL($bom, r, cStatus), "Short")
OP         ELSE
OP           SET_VALUE(CELL($bom, r, cStatus), "Check")
OP         LET(n, n + 1)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the output sheet with the Short BOM lines, sorted by Part No A to Z, and return how many there are.
PROC     ListShortParts
KIND     Function
MODULE   modSteps
ARGS     bom As Worksheet, out As Worksheet
RET      Long
VAR      cPart As Long
VAR      cReq As Long
VAR      cRec As Long
VAR      cStatus As Long
VAR      outRow As Long
VAR      n As Long
PRE      COL($bom, "Part No") > 0; COL($bom, "Required Qty") > 0; COL($bom, "Received") > 0; COL($bom, "Status") > 0
MODIFIES $out
OP       LET(cPart, COL($bom, "Part No"))
OP       LET(cReq, COL($bom, "Required Qty"))
OP       LET(cRec, COL($bom, "Received"))
OP       LET(cStatus, COL($bom, "Status"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Part No", "Required Qty", "Received", "Status")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($bom, cPart) var=r
OP         SKIP_IF IS_EMPTY(CELL($bom, r, cPart))
OP         IF SAME(CELL($bom, r, cStatus), "Short")
OP           WRITE_ROW($out, outRow, TRIM(CELL($bom, r, cPart)), CELL($bom, r, cReq), CELL($bom, r, cRec), "Short")
OP           LET(outRow, outRow + 1)
OP           LET(n, n + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       RETURN(n)
ERR      standard
