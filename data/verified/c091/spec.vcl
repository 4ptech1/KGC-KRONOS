VCL 0.6
BUILD
# ASSUME: Material Certs headers in row 1 include "Heat No", "MTC No", "Supplier", "Grade" and "Remark"
# ASSUME: the real sheets are "Material Certs", "Grade Summary" and "Chase List" and all three exist; the TESTs use Certs, Summary and Chase
# ASSUME: the rows run down to the last Heat No; every row down to there is a certificate row, even one whose Heat No is blank
# ASSUME: a Heat No is rewritten only when trimming and upper case change its text, so a number or an error value stays as it is
# ASSUME: Grade Summary and Chase List are cleared and rebuilt on every run; the No grade row is always written, last, after the grades
# ASSUME: grades and suppliers are the same when they match ignoring case and spaces at the ends; the first spelling is written
# ASSUME: a Check row with no Supplier is marked Check but adds nothing to the Chase List
GOAL     Tidy the Heat No column of the Material Certs sheet, mark rows with no MTC No or no Grade as Check, and rebuild a grade summary and a list of suppliers to chase.
NAME     MaterialCertTool
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(certs As Worksheet, summary As Worksheet, chase As Worksheet) As Long
SIG      modSteps.CleanHeatNumbers(certs As Worksheet)
SIG      modSteps.FlagIncompleteCerts(certs As Worksheet) As Long
SIG      modSteps.WriteGradeSummary(certs As Worksheet, summary As Worksheet)
SIG      modSteps.WriteChaseList(certs As Worksheet, chase As Worksheet)
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.CleanHeatNumbers
FLOW     modMain.RunHeadless -> modSteps.FlagIncompleteCerts
FLOW     modMain.RunHeadless -> modSteps.WriteGradeSummary
FLOW     modMain.RunHeadless -> modSteps.WriteChaseList
FIXTURE  certs_01.xlsx sheet=Certs cols=A:Heat No,B:MTC No,C:Supplier,D:Grade,E:Remark rows=12 data=A:values(" h100 "x1,"H200"x1,"h300"x1,"H400"x1,"H500 "x1,"H600"x1,""x1,"H800"x1,"h900"x1,"H910"x1,"H920"x1," h930"x1)+boundary(" h100 ","H200","h300","H400","H500 ","H600","","H800","h900","H910","H920"," h930");B:values("MTC-3"x1,"MTC-4"x1,"MTC-6"x1,"MTC-7"x1,"MTC-9"x1,"MTC-10"x1,"MTC-11"x1,""x5)+boundary("","","MTC-3","MTC-4","","MTC-6","MTC-7","","MTC-9","MTC-10","MTC-11","");C:values("Zenith Steel"x2,"Apex Metals"x1,"apex metals "x1,"Borealis Alloys"x2,"Corvid Tubes"x2,"Dune Forge"x1,"Ember Castings"x1,""x2)+boundary("Zenith Steel","Apex Metals","apex metals ","Borealis Alloys","","Corvid Tubes","Zenith Steel","Dune Forge","Corvid Tubes","","Borealis Alloys","Ember Castings");D:values("SA-516 Gr 70"x3,"sa-516 gr 70 "x1,"SA-240 304"x1,"SA-106 B"x5,""x2)+boundary("SA-516 Gr 70","SA-516 Gr 70","","sa-516 gr 70 ","SA-240 304","SA-516 Gr 70","SA-106 B","SA-106 B","SA-106 B","SA-106 B","","SA-106 B");E:values("old"x5,""x5,"Check"x2)+boundary("old","","Check","Check","old","","old","","old","","old","")
FIXTURE  certs_01.xlsx sheet=Summary cols=A:Grade,B:Certificates rows=0 stale=7
FIXTURE  certs_01.xlsx sheet=Chase cols=A:Supplier rows=0 stale=8
FIXTURE  certs_02.xlsx sheet=Certs cols=A:Heat No,B:MTC No,C:Supplier,D:Grade,E:Remark rows=4 data=A:values("H100"x1,"H200"x1,"H300"x1,"H400"x1)+boundary("H100","H200","H300","H400");B:values("MTC-1"x1,"MTC-2"x1,"MTC-3"x1,"MTC-4"x1)+boundary("MTC-1","MTC-2","MTC-3","MTC-4");C:values("Zenith Steel"x2,""x1,"Apex Metals"x1)+boundary("Zenith Steel","","Apex Metals","Zenith Steel");D:values("SA-516 Gr 70"x3,"SA-106 B"x1)+boundary("SA-516 Gr 70","SA-106 B","SA-516 Gr 70","SA-516 Gr 70");E:values("Check"x2,"old"x1,""x1)+boundary("Check","old","","Check")
FIXTURE  certs_02.xlsx sheet=Summary cols=A:Grade,B:Certificates rows=0 stale=5
FIXTURE  certs_02.xlsx sheet=Chase cols=A:Supplier rows=0 stale=3
TEST     fixture=certs_01.xlsx call=modMain.RunHeadless(Certs, Summary, Chase) runs=2 assert=RESULT = 7 AND ERROR_COUNT = 0 AND COUNTIF(Certs![Remark], "Check") = 7 AND COUNTIF(Certs![Remark], "old") = 0 AND COUNTA(Certs![Remark]) = 8 AND VALUE(Certs![Remark]2) = "Check" AND VALUE(Certs![Remark]4) = "Check" AND VALUE(Certs![Remark]13) = "Check" AND VALUE(Certs![Heat No]2) = "H100" AND VALUE(Certs![Heat No]4) = "H300" AND VALUE(Certs![Heat No]6) = "H500" AND VALUE(Certs![Heat No]13) = "H930" AND ROWS_USED(Summary) = 5 AND VALUE(Summary![Grade]2) = "SA-106 B" AND VALUE(Summary![Certificates]2) = 5 AND VALUE(Summary![Grade]3) = "SA-240 304" AND VALUE(Summary![Certificates]3) = 1 AND VALUE(Summary![Grade]4) = "SA-516 Gr 70" AND VALUE(Summary![Certificates]4) = 4 AND VALUE(Summary![Grade]5) = "No grade" AND VALUE(Summary![Certificates]5) = 2 AND COUNTIF(Summary![Grade], "stale") = 0
TEST     fixture=certs_01.xlsx call=modMain.RunHeadless(Certs, Summary, Chase) runs=2 assert=RESULT = 7 AND ERROR_COUNT = 0 AND ROWS_USED(Chase) = 6 AND VALUE(Chase![Supplier]2) = "Apex Metals" AND VALUE(Chase![Supplier]3) = "Borealis Alloys" AND VALUE(Chase![Supplier]4) = "Dune Forge" AND VALUE(Chase![Supplier]5) = "Ember Castings" AND VALUE(Chase![Supplier]6) = "Zenith Steel" AND COUNTIF(Chase![Supplier], "stale") = 0 AND FILL_COUNT(Chase!A1:C10, #FF0000) = 0 AND FILL_COUNT(Summary!A1:C10, #FF0000) = 0
TEST     fixture=certs_01.xlsx call=modMain.RunHeadless(Certs, Summary, Chase) assert=RESULT = 7 AND COUNTA(Certs![Heat No]) = PRE_COUNTA(Certs![Heat No]) AND COUNTA(Certs![MTC No]) = PRE_COUNTA(Certs![MTC No]) AND COUNTA(Certs![Supplier]) = PRE_COUNTA(Certs![Supplier]) AND COUNTA(Certs![Grade]) = PRE_COUNTA(Certs![Grade]) AND COUNTIF(Certs![Supplier], "Corvid Tubes") = PRE_COUNTIF(Certs![Supplier], "Corvid Tubes") AND SUM(Summary![Certificates]) = 12
TEST     fixture=certs_02.xlsx call=modMain.RunHeadless(Certs, Summary, Chase) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTA(Certs![Remark]) = 1 AND COUNTA(Certs![Heat No]) = PRE_COUNTA(Certs![Heat No]) AND ROWS_USED(Summary) = 4 AND VALUE(Summary![Grade]2) = "SA-106 B" AND VALUE(Summary![Certificates]2) = 1 AND VALUE(Summary![Grade]3) = "SA-516 Gr 70" AND VALUE(Summary![Certificates]3) = 3 AND VALUE(Summary![Grade]4) = "No grade" AND VALUE(Summary![Certificates]4) = 0 AND ROWS_USED(Chase) = 1 AND VALUE(Chase![Supplier]1) = "Supplier"
TEST     fixture=certs_01.xlsx call=modMain.RunHeadless(Summary, Certs, Chase) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.6
TASK
GOAL     Start the material certificate tool on the workbook's own sheets and tell the user how many rows are marked Check.
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "Material Certs"), SHEET(THIS_BOOK, "Grade Summary"), SHEET(THIS_BOOK, "Chase List"), into=n)
OP       MSG("Rows marked Check: " & n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Run the four steps of the material certificate tool on the three sheets it is given and return how many rows are marked Check.
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     certs As Worksheet, summary As Worksheet, chase As Worksheet
RET      Long
VAR      flagged As Long
PRE      COL($certs, "Heat No") > 0; COL($certs, "MTC No") > 0; COL($certs, "Supplier") > 0; COL($certs, "Grade") > 0; COL($certs, "Remark") > 0
MODIFIES $certs["Heat No", "Remark"]
MODIFIES $summary
MODIFIES $chase
OP       CALL(modSteps.CleanHeatNumbers, $certs)
OP       CALL(modSteps.FlagIncompleteCerts, $certs, into=flagged)
OP       CALL(modSteps.WriteGradeSummary, $certs, $summary)
OP       CALL(modSteps.WriteChaseList, $certs, $chase)
OP       RETURN(flagged)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rewrite every Heat No of the certificate sheet trimmed and in upper case.
PROC     CleanHeatNumbers
KIND     Sub
MODULE   modSteps
ARGS     certs As Worksheet
VAR      cHeat As Long
PRE      COL($certs, "Heat No") > 0
MODIFIES $certs["Heat No"]
OP       LET(cHeat, COL($certs, "Heat No"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($certs, cHeat) var=r
OP         IF CELL($certs, r, cHeat) <> UPPER(TRIM(CELL($certs, r, cHeat))) THEN SET_VALUE(CELL($certs, r, cHeat), UPPER(TRIM(CELL($certs, r, cHeat))))
ERR      standard
---
VCL 0.6
TASK
GOAL     Write Check in the Remark of every certificate row with no MTC No or no Grade, clear the Remark of the other rows and return how many rows are Check.
PROC     FlagIncompleteCerts
KIND     Function
MODULE   modSteps
ARGS     certs As Worksheet
RET      Long
VAR      cHeat As Long
VAR      cMtc As Long
VAR      cGrade As Long
VAR      cRemark As Long
VAR      n As Long
PRE      COL($certs, "Heat No") > 0; COL($certs, "MTC No") > 0; COL($certs, "Grade") > 0; COL($certs, "Remark") > 0
MODIFIES $certs["Remark"]
OP       LET(cHeat, COL($certs, "Heat No"))
OP       LET(cMtc, COL($certs, "MTC No"))
OP       LET(cGrade, COL($certs, "Grade"))
OP       LET(cRemark, COL($certs, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($certs, cHeat) var=r
OP         IF IS_EMPTY(CELL($certs, r, cMtc)) OR IS_EMPTY(CELL($certs, r, cGrade))
OP           SET_VALUE(CELL($certs, r, cRemark), "Check")
OP           LET(n, n + 1)
OP         ELSE
OP           CLEAR(CELL($certs, r, cRemark), what=contents)
OP       RETURN(n)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the grade summary sheet with one row for each Grade and its number of certificates, grades A to Z, then a No grade row.
PROC     WriteGradeSummary
KIND     Sub
MODULE   modSteps
ARGS     certs As Worksheet, summary As Worksheet
VAR      cHeat As Long
VAR      cGrade As Long
VAR      counts As Dictionary
VAR      got As Long
VAR      noGrade As Long
VAR      outRow As Long
PRE      COL($certs, "Heat No") > 0; COL($certs, "Grade") > 0
MODIFIES $summary
OP       LET(cHeat, COL($certs, "Heat No"))
OP       LET(cGrade, COL($certs, "Grade"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($certs, cHeat) var=r
OP         IF IS_EMPTY(CELL($certs, r, cGrade))
OP           LET(noGrade, noGrade + 1)
OP         ELSE
OP           DICT_ADD(counts, TRIM(CELL($certs, r, cGrade)), 1, mode=sum)
OP       CLEAR(USED_RANGE($summary), what=all)
OP       WRITE_ROW($summary, 1, "Grade", "Certificates")
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         WRITE_ROW($summary, outRow, k, got)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($summary), key=1, order=asc, header=yes)
OP       WRITE_ROW($summary, outRow, "No grade", noGrade)
ERR      standard
---
VCL 0.6
TASK
GOAL     Rebuild the chase list sheet with each Supplier that has a certificate row with no MTC No or no Grade, written once, A to Z.
PROC     WriteChaseList
KIND     Sub
MODULE   modSteps
ARGS     certs As Worksheet, chase As Worksheet
VAR      cHeat As Long
VAR      cMtc As Long
VAR      cSupplier As Long
VAR      cGrade As Long
VAR      seen As Dictionary
VAR      outRow As Long
PRE      COL($certs, "Heat No") > 0; COL($certs, "MTC No") > 0; COL($certs, "Supplier") > 0; COL($certs, "Grade") > 0
MODIFIES $chase
OP       LET(cHeat, COL($certs, "Heat No"))
OP       LET(cMtc, COL($certs, "MTC No"))
OP       LET(cSupplier, COL($certs, "Supplier"))
OP       LET(cGrade, COL($certs, "Grade"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($certs, cHeat) var=r
OP         SKIP_IF IS_EMPTY(CELL($certs, r, cSupplier))
OP         IF IS_EMPTY(CELL($certs, r, cMtc)) OR IS_EMPTY(CELL($certs, r, cGrade)) THEN DICT_ADD(seen, TRIM(CELL($certs, r, cSupplier)), 1, mode=keep)
OP       CLEAR(USED_RANGE($chase), what=all)
OP       WRITE_ROW($chase, 1, "Supplier")
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(seen)
OP         WRITE_ROW($chase, outRow, k)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($chase), key=1, order=asc, header=yes)
ERR      standard
