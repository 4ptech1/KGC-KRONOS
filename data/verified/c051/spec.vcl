VCL 0.6
BUILD
# ASSUME: GRN headers (row 1) include "GRN No", "Part No", "Heat No", "Received Qty" and "Remark"; Material Certs has "Heat No" and "MTC No"; BOM has "Part No" and "Required Qty"
# ASSUME: the Remark column already exists on GRN; on a good line any old Remark is cleared, as asked
# ASSUME: a GRN line is good when Received Qty is a number above 0 and Heat No is not blank; any other line gets Check
# ASSUME: Missing Certs, Shortages and Receiving Summary already exist and are rebuilt from scratch on every run
# ASSUME: Missing Certs lists GRN No, Part No and Heat No of each GRN line whose non-blank Heat No is not in Material Certs; a blank Heat No is already Check, so it is not listed
# ASSUME: a heat is in Material Certs when it is in its Heat No column, ignoring case and spaces; MTC No is not checked
# ASSUME: Shortages lists each BOM line whose Required Qty is a number above the total Received Qty of the good GRN lines of the same Part No (ignoring case and spaces); columns are Part No, Required Qty, Valid Qty and Shortfall
# ASSUME: a GRN line with a blank Part No is still checked and listed in Missing Certs, but adds to no BOM part
# ASSUME: a BOM line with a blank Part No or a Required Qty that is not a number is skipped; each BOM line is compared on its own
# ASSUME: Receiving Summary has the headers Measure and Count, then one row each for Check lines, Missing certs and Short parts
# ASSUME: the real sheets are GRN, Material Certs, BOM, Missing Certs, Shortages and Receiving Summary; the TESTs use GRN, Certs, BOM, MissingCerts, Shortages and RecSummary
GOAL     Flag bad GRN lines, list GRN lines with no material cert and BOM parts short of good received quantity, and write a receiving summary.
NAME     ReceivingCheck
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modReceiving
SIG      modMain.Run()
SIG      modMain.RunHeadless(grn As Worksheet, certs As Worksheet, bom As Worksheet, missing As Worksheet, shortages As Worksheet, summary As Worksheet) As Long
SIG      modReceiving.FlagGrnLines(grn As Worksheet) As Long
SIG      modReceiving.ListMissingCerts(grn As Worksheet, certs As Worksheet, missing As Worksheet) As Long
SIG      modReceiving.ListShortages(grn As Worksheet, bom As Worksheet, shortages As Worksheet) As Long
SIG      modReceiving.WriteSummary(summary As Worksheet, checkLines As Long, noCert As Long, shortParts As Long)
FLOW     modMain.Run -> modMain.RunHeadless -> modReceiving.FlagGrnLines
FLOW     modMain.RunHeadless -> modReceiving.ListMissingCerts
FLOW     modMain.RunHeadless -> modReceiving.ListShortages
FLOW     modMain.RunHeadless -> modReceiving.WriteSummary
FIXTURE  receiving_01.xlsx sheet=GRN cols=A:GRN No,B:Part No,C:Heat No,D:Received Qty,E:Remark rows=12 data=A:seq(1,1);B:values("P-1"x2,"P-2"x3,"P-3"x2,"P-4"x1,""x1,"P-5"x1," p-1 "x1,"P-6"x1)+boundary("P-1","P-1","P-2","P-2","P-3","P-3","P-4","","P-5"," p-1 ","P-6","P-2");C:values("H-9"x3,"H-2"x3,"H-3"x1,""x1,"H-7"x1,"H-1"x1," h-1 "x1," H-8 "x1)+boundary("H-9","H-2","H-3","H-9","","H-7","H-2","H-1"," h-1 ","H-2"," H-8 ","H-9");D:int(1,50)+boundary(10,5,0,20,8,-4,"",2,3,1,4,2);E:values("old"x12)
FIXTURE  receiving_01.xlsx sheet=Certs cols=A:Heat No,B:MTC No rows=5 data=A:values(" H-3 "x1,"H-2"x2,"H-4"x1,"H-1"x1)+boundary(" H-3 ","H-2","H-4","H-2","H-1");B:seq(100,1)
FIXTURE  receiving_01.xlsx sheet=BOM cols=A:Part No,B:Required Qty rows=10 data=A:values("P-3"x1,"P-1"x1,"P-2"x1,"P-4"x1,"P-5"x1,"P-6"x1,""x1,"P-8"x1,"P-10"x1,"P-7"x1)+boundary("P-3","P-1","P-2","P-4","P-5","P-6","","P-8","P-10","P-7");B:int(0,30)+boundary(10,16,23,5,3,5,4,"",0,6)
FIXTURE  receiving_01.xlsx sheet=MissingCerts cols=A:GRN No,B:Part No,C:Heat No rows=0 stale=8
FIXTURE  receiving_01.xlsx sheet=Shortages cols=A:Part No,B:Required Qty,C:Valid Qty,D:Shortfall rows=0 stale=8
FIXTURE  receiving_01.xlsx sheet=RecSummary cols=A:Measure,B:Count rows=0 stale=6
FIXTURE  receiving_02.xlsx sheet=GRN cols=A:GRN No,B:Part No,C:Heat No,D:Received Qty,E:Remark rows=12 data=A:seq(1,1);B:values("P-1"x2,"P-2"x3,"P-3"x2,"P-4"x1,""x1,"P-5"x1," p-1 "x1,"P-6"x1)+boundary("P-1","P-1","P-2","P-2","P-3","P-3","P-4","","P-5"," p-1 ","P-6","P-2");C:values("H-9"x3,"H-2"x3,"H-3"x1,""x1,"H-7"x1,"H-1"x1," h-1 "x1," H-8 "x1)+boundary("H-9","H-2","H-3","H-9","","H-7","H-2","H-1"," h-1 ","H-2"," H-8 ","H-9");D:values("10"x1,"5"x1,"0"x1,"20"x1,"8"x1,"-4"x1,"abc"x1,"2"x2,"3"x1,"1"x1,"4"x1)+boundary("10","5","0","20","8","-4","abc","2","3","1","4","2");E:values("old"x12)
FIXTURE  receiving_02.xlsx sheet=Certs cols=A:Heat No,B:MTC No rows=5 data=A:values(" H-3 "x1,"H-2"x2,"H-4"x1,"H-1"x1)+boundary(" H-3 ","H-2","H-4","H-2","H-1");B:seq(100,1)
FIXTURE  receiving_02.xlsx sheet=BOM cols=A:Part No,B:Required Qty rows=10 data=A:values("P-3"x1,"P-1"x1,"P-2"x1,"P-4"x1,"P-5"x1,"P-6"x1,""x1,"P-8"x1,"P-10"x1,"P-7"x1)+boundary("P-3","P-1","P-2","P-4","P-5","P-6","","P-8","P-10","P-7");B:values("10"x1,"16"x1,"23"x1,"5"x2,"3"x1,"4"x1,"n/a"x1,"0"x1,"6"x1)+boundary("10","16","23","5","3","5","4","n/a","0","6")
FIXTURE  receiving_02.xlsx sheet=MissingCerts cols=A:GRN No,B:Part No,C:Heat No rows=0 stale=8
FIXTURE  receiving_02.xlsx sheet=Shortages cols=A:Part No,B:Required Qty,C:Valid Qty,D:Shortfall rows=0 stale=8
FIXTURE  receiving_02.xlsx sheet=RecSummary cols=A:Measure,B:Count rows=0 stale=6
TEST     fixture=receiving_01.xlsx call=modMain.RunHeadless(GRN, Certs, BOM, MissingCerts, Shortages, RecSummary) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(GRN![Remark], "Check") = 4 AND COUNTIF(GRN![Remark], "old") = 0 AND COUNTA(GRN![Remark]) = 5 AND COUNTA(MissingCerts![GRN No]) = 6 AND COUNTA(Shortages![Part No]) = 6 AND COUNTA(RecSummary![Measure]) = 4 AND COUNTIF(MissingCerts![GRN No], "stale") = 0 AND COUNTIF(Shortages![Part No], "stale") = 0 AND COUNTIF(RecSummary![Measure], "stale") = 0 AND FILL_COUNT(MissingCerts![GRN No], #FF0000) = 0 AND FILL_COUNT(Shortages![Part No], #FF0000) = 0 AND FILL_COUNT(RecSummary![Measure], #FF0000) = 0
TEST     fixture=receiving_01.xlsx call=modMain.RunHeadless(GRN, Certs, BOM, MissingCerts, Shortages, RecSummary) assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(GRN![GRN No]) = PRE_COUNTA(GRN![GRN No]) AND COUNTA(GRN![Part No]) = PRE_COUNTA(GRN![Part No]) AND COUNTA(GRN![Heat No]) = PRE_COUNTA(GRN![Heat No]) AND COUNTA(GRN![Received Qty]) = PRE_COUNTA(GRN![Received Qty]) AND COUNTA(Certs![Heat No]) = PRE_COUNTA(Certs![Heat No]) AND COUNTA(BOM![Part No]) = PRE_COUNTA(BOM![Part No]) AND COUNTA(BOM![Required Qty]) = PRE_COUNTA(BOM![Required Qty]) AND VALUE(GRN![Remark]4) = "Check" AND VALUE(GRN![Remark]6) = "Check" AND VALUE(GRN![Remark]7) = "Check" AND VALUE(GRN![Remark]8) = "Check" AND VALUE(GRN![Remark]2) = "" AND VALUE(GRN![Remark]13) = ""
TEST     fixture=receiving_01.xlsx call=modMain.RunHeadless(GRN, Certs, BOM, MissingCerts, Shortages, RecSummary) assert=ROWS_USED(MissingCerts) = 6 AND COUNTIF(MissingCerts![Heat No], "H-9") = 3 AND COUNTIF(MissingCerts![Heat No], "H-7") = 1 AND COUNTIF(MissingCerts![Heat No], "H-8") = 1 AND VALUE(MissingCerts![GRN No]2) = 1 AND VALUE(MissingCerts![GRN No]3) = 4 AND VALUE(MissingCerts![GRN No]4) = 6 AND VALUE(MissingCerts![GRN No]5) = 11 AND VALUE(MissingCerts![GRN No]6) = 12 AND VALUE(MissingCerts![Part No]2) = "P-1" AND VALUE(MissingCerts![Part No]4) = "P-3" AND VALUE(MissingCerts![Part No]5) = "P-6" AND VALUE(MissingCerts![Heat No]5) = "H-8"
TEST     fixture=receiving_01.xlsx call=modMain.RunHeadless(GRN, Certs, BOM, MissingCerts, Shortages, RecSummary) assert=ROWS_USED(Shortages) = 6 AND SUM(Shortages![Shortfall]) = 23 AND VALUE(Shortages![Part No]2) = "P-3" AND VALUE(Shortages![Required Qty]2) = 10 AND VALUE(Shortages![Valid Qty]2) = 0 AND VALUE(Shortages![Shortfall]2) = 10 AND VALUE(Shortages![Part No]3) = "P-2" AND VALUE(Shortages![Valid Qty]3) = 22 AND VALUE(Shortages![Shortfall]3) = 1 AND VALUE(Shortages![Part No]4) = "P-4" AND VALUE(Shortages![Shortfall]4) = 5 AND VALUE(Shortages![Part No]5) = "P-6" AND VALUE(Shortages![Valid Qty]5) = 4 AND VALUE(Shortages![Part No]6) = "P-7" AND VALUE(Shortages![Shortfall]6) = 6
TEST     fixture=receiving_01.xlsx call=modMain.RunHeadless(GRN, Certs, BOM, MissingCerts, Shortages, RecSummary) assert=ROWS_USED(RecSummary) = 4 AND VALUE(RecSummary![Measure]2) = "Check lines" AND VALUE(RecSummary![Count]2) = 4 AND VALUE(RecSummary![Measure]3) = "Missing certs" AND VALUE(RecSummary![Count]3) = 5 AND VALUE(RecSummary![Measure]4) = "Short parts" AND VALUE(RecSummary![Count]4) = 5
TEST     fixture=receiving_02.xlsx call=modMain.RunHeadless(GRN, Certs, BOM, MissingCerts, Shortages, RecSummary) assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(GRN![Remark], "Check") = 4 AND VALUE(GRN![Remark]8) = "Check" AND ROWS_USED(MissingCerts) = 6 AND ROWS_USED(Shortages) = 6 AND SUM(Shortages![Shortfall]) = 23 AND VALUE(RecSummary![Count]3) = 5
TEST     fixture=receiving_01.xlsx call=modMain.RunHeadless(Certs, GRN, BOM, MissingCerts, Shortages, RecSummary) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
TASK
GOAL     Start the receiving check on the workbook's own sheets and tell the user how many GRN lines were flagged Check.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "GRN"), SHEET(THIS_BOOK, "Material Certs"), SHEET(THIS_BOOK, "BOM"), SHEET(THIS_BOOK, "Missing Certs"), SHEET(THIS_BOOK, "Shortages"), SHEET(THIS_BOOK, "Receiving Summary"), into=n)
OP       MSG("GRN lines flagged Check: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run all four receiving steps on the sheets it is given and return how many GRN lines were flagged Check.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     grn As Worksheet, certs As Worksheet, bom As Worksheet, missing As Worksheet, shortages As Worksheet, summary As Worksheet
RET      Long
VAR      nCheck As Long
VAR      nMissing As Long
VAR      nShort As Long
PRE      COL($grn, "GRN No") > 0; COL($grn, "Part No") > 0; COL($grn, "Heat No") > 0; COL($grn, "Received Qty") > 0; COL($grn, "Remark") > 0; COL($certs, "Heat No") > 0; COL($bom, "Part No") > 0; COL($bom, "Required Qty") > 0
MODIFIES $grn["Remark"]
MODIFIES $missing
MODIFIES $shortages
MODIFIES $summary
OP       CALL(modReceiving.FlagGrnLines, $grn, into=nCheck)
OP       CALL(modReceiving.ListMissingCerts, $grn, $certs, $missing, into=nMissing)
OP       CALL(modReceiving.ListShortages, $grn, $bom, $shortages, into=nShort)
OP       CALL(modReceiving.WriteSummary, $summary, nCheck, nMissing, nShort)
OP       RETURN(nCheck)
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Check in the Remark column of every GRN line with no positive received quantity or no heat number, clear it on the others and return how many were flagged.
PROC     FlagGrnLines
KIND     Function
MODULE   modReceiving
ARGS     grn As Worksheet
RET      Long
VAR      cGrn As Long
VAR      cHeat As Long
VAR      cQty As Long
VAR      cRemark As Long
VAR      isBad As Boolean
VAR      n As Long
PRE      COL($grn, "GRN No") > 0; COL($grn, "Heat No") > 0; COL($grn, "Received Qty") > 0; COL($grn, "Remark") > 0
MODIFIES $grn["Remark"]
OP       LET(cGrn, COL($grn, "GRN No"))
OP       LET(cHeat, COL($grn, "Heat No"))
OP       LET(cQty, COL($grn, "Received Qty"))
OP       LET(cRemark, COL($grn, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($grn, cGrn) var=r
OP         LET(isBad, FALSE)
OP         IF IS_NUMBER(CELL($grn, r, cQty))
OP           IF CELL($grn, r, cQty) <= 0 THEN LET(isBad, TRUE)
OP         ELSE
OP           LET(isBad, TRUE)
OP         IF IS_EMPTY(CELL($grn, r, cHeat)) THEN LET(isBad, TRUE)
OP         IF isBad
OP           SET_VALUE(CELL($grn, r, cRemark), "Check")
OP           LET(n, n + 1)
OP         ELSE
OP           CLEAR(CELL($grn, r, cRemark), what=contents)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Missing Certs sheet with the GRN lines whose heat number is not in Material Certs and return how many were listed.
PROC     ListMissingCerts
KIND     Function
MODULE   modReceiving
ARGS     grn As Worksheet, certs As Worksheet, missing As Worksheet
RET      Long
VAR      cGrn As Long
VAR      cPart As Long
VAR      cHeat As Long
VAR      cCert As Long
VAR      known As Dictionary
VAR      outRow As Long
PRE      COL($grn, "GRN No") > 0; COL($grn, "Part No") > 0; COL($grn, "Heat No") > 0; COL($certs, "Heat No") > 0
MODIFIES $missing
OP       LET(cGrn, COL($grn, "GRN No"))
OP       LET(cPart, COL($grn, "Part No"))
OP       LET(cHeat, COL($grn, "Heat No"))
OP       LET(cCert, COL($certs, "Heat No"))
OP       CLEAR(USED_RANGE($missing), what=all)
OP       WRITE_ROW($missing, 1, "GRN No", "Part No", "Heat No")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($certs, cCert) var=r
OP         DICT_ADD(known, TRIM(CELL($certs, r, cCert)), 1, mode=keep)
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($grn, cGrn) var=r
OP         SKIP_IF IS_EMPTY(CELL($grn, r, cHeat))
OP         SKIP_IF EXISTS(known, TRIM(CELL($grn, r, cHeat)))
OP         WRITE_ROW($missing, outRow, CELL($grn, r, cGrn), CELL($grn, r, cPart), TRIM(CELL($grn, r, cHeat)))
OP         LET(outRow, outRow + 1)
OP       RETURN(outRow - 2)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Shortages sheet with each BOM part whose good received quantity is below the required quantity, with the shortfall, and return how many parts were listed.
PROC     ListShortages
KIND     Function
MODULE   modReceiving
ARGS     grn As Worksheet, bom As Worksheet, shortages As Worksheet
RET      Long
VAR      cGrnPart As Long
VAR      cHeat As Long
VAR      cQty As Long
VAR      cBomPart As Long
VAR      cReq As Long
VAR      totals As Dictionary
VAR      got As Double
VAR      outRow As Long
PRE      COL($grn, "Part No") > 0; COL($grn, "Heat No") > 0; COL($grn, "Received Qty") > 0; COL($bom, "Part No") > 0; COL($bom, "Required Qty") > 0
MODIFIES $shortages
OP       LET(cGrnPart, COL($grn, "Part No"))
OP       LET(cHeat, COL($grn, "Heat No"))
OP       LET(cQty, COL($grn, "Received Qty"))
OP       LET(cBomPart, COL($bom, "Part No"))
OP       LET(cReq, COL($bom, "Required Qty"))
OP       CLEAR(USED_RANGE($shortages), what=all)
OP       WRITE_ROW($shortages, 1, "Part No", "Required Qty", "Valid Qty", "Shortfall")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($grn, cGrnPart) var=r
OP         SKIP_IF IS_EMPTY(CELL($grn, r, cHeat))
OP         SKIP_IF NOT IS_NUMBER(CELL($grn, r, cQty))
OP         IF CELL($grn, r, cQty) > 0 THEN DICT_ADD(totals, TRIM(CELL($grn, r, cGrnPart)), CELL($grn, r, cQty), mode=sum)
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($bom, cBomPart) var=r
OP         SKIP_IF IS_EMPTY(CELL($bom, r, cBomPart))
OP         SKIP_IF NOT IS_NUMBER(CELL($bom, r, cReq))
OP         LET(got, 0)
OP         IF EXISTS(totals, TRIM(CELL($bom, r, cBomPart)))
OP           DICT_GET(totals, TRIM(CELL($bom, r, cBomPart)), into=got)
OP         IF CELL($bom, r, cReq) > got
OP           WRITE_ROW($shortages, outRow, TRIM(CELL($bom, r, cBomPart)), CELL($bom, r, cReq), got, CELL($bom, r, cReq) - got)
OP           LET(outRow, outRow + 1)
OP       RETURN(outRow - 2)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the Receiving Summary sheet with the number of Check lines, missing certs and short parts it is given.
PROC     WriteSummary
KIND     Sub
MODULE   modReceiving
ARGS     summary As Worksheet, checkLines As Long, noCert As Long, shortParts As Long
MODIFIES $summary
OP       CLEAR(USED_RANGE($summary), what=all)
OP       WRITE_ROW($summary, 1, "Measure", "Count")
OP       WRITE_ROW($summary, 2, "Check lines", $checkLines)
OP       WRITE_ROW($summary, 3, "Missing certs", $noCert)
OP       WRITE_ROW($summary, 4, "Short parts", $shortParts)
ERR      standard
