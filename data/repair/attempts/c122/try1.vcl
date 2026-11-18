VCL 0.6
TASK
# ASSUME: Material Certs headers in row 1 include "Heat No" and "MTC No"; GRN headers include "GRN No" and "Heat No"; the real sheets are "Material Certs" and "GRN"
# ASSUME: Heat numbers are matched with CLEAN_TEXT, ignoring case; a cert row with no Heat No is left as it is
# ASSUME: each GRN line with a Heat No counts once; GRN lines with no Heat No are ignored
# ASSUME: the GRN Check heading is added after the last column of Material Certs when missing
GOAL     Write Not received or the number of GRN lines in GRN Check for each heat on Material Certs, and return how many heats were not received.
PROC     CheckHeatsAgainstGrn
KIND     Function
MODULE   modCerts
ARGS     certs As Worksheet, grn As Worksheet
RET      Long
VAR      cHeat As Long
VAR      cCheck As Long
VAR      gHeat As Long
VAR      counts As Dictionary
VAR      lines As Long
VAR      missing As Long
PRE      COL($certs, "Heat No") > 0; COL($certs, "MTC No") > 0; COL($grn, "GRN No") > 0; COL($grn, "Heat No") > 0
MODIFIES $certs
OP       LET(cHeat, COL($certs, "Heat No"))
OP       LET(gHeat, COL($grn, "Heat No"))
OP       LET(cCheck, COL($certs, "GRN Check"))
OP       IF cCheck = 0
OP         LET(cCheck, LASTCOL($certs, 1) + 1)
OP         SET_VALUE(CELL($certs, 1, cCheck), "GRN Check")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($grn, gHeat) var=g
OP         SKIP_IF IS_EMPTY(CELL($grn, g, gHeat))
OP         DICT_ADD(counts, CLEAN_TEXT(CELL($grn, g, gHeat)), 1, mode=sum)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($certs, cHeat) var=r
OP         SKIP_IF IS_EMPTY(CELL($certs, r, cHeat))
OP         IF EXISTS(counts, CLEAN_TEXT(CELL($certs, r, cHeat)))
OP           DICT_GET(counts, CLEAN_TEXT(CELL($certs, r, cHeat)), into=lines)
OP           SET_VALUE(CELL($certs, r, cCheck), lines)
OP         ELSE
OP           SET_VALUE(CELL($certs, r, cCheck), "Not received")
OP           LET(missing, missing + 1)
OP       RETURN(missing)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  certs_01.xlsx sheet=Certs cols=A:Heat No,B:MTC No,C:GRN Check rows=8 stale=2 data=A:values("H-100"x1,"h-101 "x1,"H-102"x1,""x1,"H-103"x1,"H-104"x1,"h-100"x1,"H-105"x1)+boundary("H-100","h-101 ","H-102","","H-103","H-104","h-100","H-105");B:text(8);C:values("old"x7,"keep"x1)+boundary("old","old","old","keep")
FIXTURE  certs_01.xlsx sheet=Grn cols=A:GRN No,B:Part No,C:Heat No rows=8 stale=1 data=A:seq(1,1);B:text(6);C:values("H-100"x1," h-100"x1,"H-101"x1,"H-103"x2,"h-103"x1,""x1,"H-199"x1)+boundary("H-100"," h-100","H-101","H-103","h-103","H-103","","H-199")
FIXTURE  certs_02.xlsx sheet=Certs cols=A:Heat No,B:MTC No rows=4 stale=1 data=A:values("H-100"x2,""x1,"H-200"x1)+boundary("H-100","","H-200","H-100");B:text(8)
FIXTURE  certs_02.xlsx sheet=Grn cols=A:GRN No,B:Part No,C:Heat No rows=3 stale=1 data=A:seq(1,1);B:text(6);C:values("H-100"x2,""x1)+boundary("H-100","","H-100")
TEST     fixture=certs_01.xlsx call=CheckHeatsAgainstGrn(Certs, Grn) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Certs![GRN Check], "Not received") = 3 AND COUNTIF(Certs![GRN Check], "old") = 0 AND COUNTIF(Certs![GRN Check], "keep") = 1 AND VALUE(Certs![GRN Check]2) = 2 AND VALUE(Certs![GRN Check]3) = 1 AND VALUE(Certs![GRN Check]4) = "Not received" AND VALUE(Certs![GRN Check]6) = 3 AND VALUE(Certs![GRN Check]8) = 2
TEST     fixture=certs_01.xlsx call=CheckHeatsAgainstGrn(Certs, Grn) assert=COUNTA(Certs![Heat No]) = PRE_COUNTA(Certs![Heat No]) AND COUNTA(Certs![MTC No]) = PRE_COUNTA(Certs![MTC No]) AND COUNTA(Grn![Heat No]) = PRE_COUNTA(Grn![Heat No]) AND COUNTA(Grn![GRN No]) = PRE_COUNTA(Grn![GRN No])
TEST     fixture=certs_02.xlsx call=CheckHeatsAgainstGrn(Certs, Grn) assert=RESULT = 1 AND ERROR_COUNT = 0 AND VALUE(Certs![GRN Check]1) = "GRN Check" AND VALUE(Certs![GRN Check]2) = 2 AND VALUE(Certs![GRN Check]4) = "Not received" AND VALUE(Certs![GRN Check]5) = 2
TEST     fixture=certs_01.xlsx call=CheckHeatsAgainstGrn(Grn, Certs) assert=RESULT = -1 AND ERROR_COUNT = 1
