VCL 0.6
TASK
# ASSUME: BOM headers in row 1 include "Part No" and "Heat No"; Material Certs headers include "Heat No" and "MTC No"
# ASSUME: when the BOM has no "MTC No" heading it is added in the first free column of row 1
# ASSUME: the BOM rows run down to the last Part No; the MTC No column belongs to this macro, so it is rewritten on every row and cleared on a line with no Heat No
# ASSUME: heats are matched ignoring case and spaces at the ends; the certificate number is copied as it is typed
# ASSUME: a certificate row with no MTC No is not a certificate, so its heat is Missing; when a heat has several certificates the first one is used
# ASSUME: the number of Missing lines is the macro's return value
GOAL     Fill the MTC No column of the BOM with the certificate number of each line's heat from the Material Certs sheet, write Missing where the heat has no certificate and return how many are Missing.
PROC     FillMtcNumbers
KIND     Function
MODULE   modCerts
ARGS     bom As Worksheet, certs As Worksheet
RET      Long
VAR      cPart As Long
VAR      cHeat As Long
VAR      cMtc As Long
VAR      kHeat As Long
VAR      kMtc As Long
VAR      byHeat As Dictionary
VAR      mtc As Variant
VAR      n As Long
PRE      COL($bom, "Part No") > 0; COL($bom, "Heat No") > 0; COL($certs, "Heat No") > 0; COL($certs, "MTC No") > 0
MODIFIES $bom
OP       LET(cPart, COL($bom, "Part No"))
OP       LET(cHeat, COL($bom, "Heat No"))
OP       LET(cMtc, COL($bom, "MTC No"))
OP       LET(kHeat, COL($certs, "Heat No"))
OP       LET(kMtc, COL($certs, "MTC No"))
OP       IF cMtc = 0
OP         LET(cMtc, LASTCOL($bom, 1) + 1)
OP         SET_VALUE(CELL($bom, 1, cMtc), "MTC No")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($certs, kHeat) var=i
OP         SKIP_IF IS_EMPTY(CELL($certs, i, kMtc))
OP         DICT_ADD(byHeat, TRIM(CELL($certs, i, kHeat)), CELL($certs, i, kMtc), mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($bom, cPart) var=r
OP         IF IS_EMPTY(CELL($bom, r, cHeat))
OP           CLEAR(CELL($bom, r, cMtc), what=contents)
OP         ELSE
OP           IF EXISTS(byHeat, TRIM(CELL($bom, r, cHeat)))
OP             DICT_GET(byHeat, TRIM(CELL($bom, r, cHeat)), into=mtc)
OP             SET_VALUE(CELL($bom, r, cMtc), mtc)
OP           ELSE
OP             SET_VALUE(CELL($bom, r, cMtc), "Missing")
OP             LET(n, n + 1)
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  mtc_01.xlsx sheet=BOM cols=A:Part No,B:Description,C:Heat No,D:MTC No,E:Remark rows=11 data=A:seq(101,1);B:text(10);C:values("H100"x1," h100 "x1,"H200"x1,""x3,"H999"x1,"h300"x1,"H400 "x1,"h600"x1,"H888"x1)+boundary("H100"," h100 ","H200","","H999","h300","H400 ","","h600","H888","");D:values("old"x5,""x4,"Missing"x1,"MTC-9"x1)+boundary("old","","Missing","old","","old","MTC-9","","old","","old");E:values("note"x3,""x8)
FIXTURE  mtc_01.xlsx sheet=Certs cols=A:Heat No,B:MTC No,C:Grade rows=8 data=A:values("H100"x1,"h200 "x1,"H300"x1,""x1,"H400"x1," H400"x1,"H500"x1,"H600"x1)+boundary("H100","h200 ","H300","","H400"," H400","H500","H600");B:values("MTC-1"x1,"MTC-2"x1,""x1,"MTC-7"x1,"MTC-4"x1,"MTC-4B"x1,"MTC-5"x1,"MTC-6"x1)+boundary("MTC-1","MTC-2","","MTC-7","MTC-4","MTC-4B","MTC-5","MTC-6");C:values("SA-516 Gr 70"x5,"SA-240 304"x3)
FIXTURE  mtc_02.xlsx sheet=BOM cols=A:Part No,B:Description,C:Heat No rows=4 data=A:seq(201,1);B:text(10);C:values("H100"x1,""x1,"H777"x1,"h200"x1)+boundary("H100","","H777","h200")
FIXTURE  mtc_02.xlsx sheet=Certs cols=A:Heat No,B:MTC No,C:Grade rows=3 data=A:values("H200"x1,"H100"x1,"H300"x1)+boundary("H200","H100","H300");B:values("MTC-2"x1,"MTC-1"x1,"MTC-3"x1)+boundary("MTC-2","MTC-1","MTC-3");C:values("SA-516 Gr 70"x3)
TEST     fixture=mtc_01.xlsx call=FillMtcNumbers(BOM, Certs) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(BOM![MTC No], "Missing") = 3 AND COUNTIF(BOM![MTC No], "old") = 0 AND COUNTA(BOM![MTC No]) = 9 AND VALUE(BOM![MTC No]2) = "MTC-1" AND VALUE(BOM![MTC No]3) = "MTC-1" AND VALUE(BOM![MTC No]4) = "MTC-2" AND VALUE(BOM![MTC No]6) = "Missing" AND VALUE(BOM![MTC No]7) = "Missing" AND VALUE(BOM![MTC No]8) = "MTC-4" AND VALUE(BOM![MTC No]10) = "MTC-6" AND VALUE(BOM![MTC No]11) = "Missing"
TEST     fixture=mtc_01.xlsx call=FillMtcNumbers(BOM, Certs) assert=RESULT = 3 AND COUNTA(BOM![Part No]) = PRE_COUNTA(BOM![Part No]) AND COUNTA(BOM![Heat No]) = PRE_COUNTA(BOM![Heat No]) AND COUNTIF(BOM![Remark], "note") = PRE_COUNTIF(BOM![Remark], "note") AND COUNTA(Certs![Heat No]) = PRE_COUNTA(Certs![Heat No]) AND COUNTA(Certs![MTC No]) = PRE_COUNTA(Certs![MTC No]) AND COUNTA(BOM!F1:F12) = 0
TEST     fixture=mtc_02.xlsx call=FillMtcNumbers(BOM, Certs) runs=2 assert=RESULT = 1 AND ERROR_COUNT = 0 AND VALUE(BOM!D1) = "MTC No" AND VALUE(BOM!D2) = "MTC-1" AND VALUE(BOM!D4) = "Missing" AND VALUE(BOM!D5) = "MTC-2" AND COUNTA(BOM!D1:D9) = 4 AND COUNTA(BOM!E1:E9) = 0 AND COUNTA(BOM![Heat No]) = PRE_COUNTA(BOM![Heat No])
TEST     fixture=mtc_01.xlsx call=FillMtcNumbers(Certs, BOM) assert=RESULT = -1 AND ERROR_COUNT = 1
