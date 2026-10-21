VCL 0.6
TASK
# ASSUME: GRN headers (row 1) include "GRN No", "Part No" and "Heat No"; Material Certs includes "Heat No"
# ASSUME: Missing Certs is an existing sheet; it is cleared and rebuilt on every run
# ASSUME: heat numbers match after trimming both ends and ignoring case; spaces inside a heat number count
# ASSUME: a GRN line with a blank Heat No has no cert row, so it is listed; a row with a blank GRN No is not a line and is skipped
# ASSUME: blank Heat No cells in Material Certs are ignored
GOAL     List every GRN line whose heat number has no row in Material Certs on the Missing Certs sheet and return how many were listed.
PROC     ListMissingCerts
KIND     Function
MODULE   modReconcile
ARGS     grn As Worksheet, certs As Worksheet, out As Worksheet
RET      Long
VAR      cGrnNo As Long
VAR      cPart As Long
VAR      cHeat As Long
VAR      cCert As Long
VAR      known As Dictionary
VAR      n As Long
PRE      COL($grn, "GRN No") > 0; COL($grn, "Part No") > 0; COL($grn, "Heat No") > 0; COL($certs, "Heat No") > 0
MODIFIES $out
OP       LET(cGrnNo, COL($grn, "GRN No"))
OP       LET(cPart, COL($grn, "Part No"))
OP       LET(cHeat, COL($grn, "Heat No"))
OP       LET(cCert, COL($certs, "Heat No"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($certs, cCert) var=r
OP         SKIP_IF IS_EMPTY(CELL($certs, r, cCert))
OP         DICT_ADD(known, TRIM(CELL($certs, r, cCert)), 1, mode=keep)
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "GRN No", "Part No", "Heat No")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($grn, cGrnNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($grn, r, cGrnNo))
OP         IF NOT EXISTS(known, TRIM(CELL($grn, r, cHeat)))
OP           LET(n, n + 1)
OP           WRITE_ROW($out, n + 1, CELL($grn, r, cGrnNo), CELL($grn, r, cPart), CELL($grn, r, cHeat))
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  grn_01.xlsx sheet=GRN cols=A:GRN No,B:Part No,C:Heat No rows=8 data=A:values("G-1"x1,"G-2"x1,"G-3"x1,""x1,"G-5"x1,"G-6"x1,"G-7"x1,"G-8"x1)+boundary("G-1","G-2","G-3","","G-5","G-6","G-7","G-8");B:values("P-1"x1,"P-2"x1,"P-3"x1,"P-4"x1,"P-5"x1,"P-6"x1,"P-7"x1,"P-8"x1)+boundary("P-1","P-2","P-3","P-4","P-5","P-6","P-7","P-8");C:values("H-300"x1," h-200 "x1,"H-100"x1,"H-999"x1,""x1,"h-100"x1," H-400 "x1,"H-500"x1)+boundary("H-300"," h-200 ","H-100","H-999","","h-100"," H-400 ","H-500")
FIXTURE  grn_01.xlsx sheet=Certs cols=A:Heat No,B:MTC No rows=6 data=A:values("H-100"x2,"H-200"x1,""x1,"H-500 "x1,"H-600"x1)+boundary("H-200","H-100","","H-100","H-500 ","H-600");B:values("M-1"x1,"M-2"x1,"M-3"x1,"M-4"x1,"M-5"x1,"M-6"x1)+boundary("M-1","M-2","M-3","M-4","M-5","M-6")
FIXTURE  grn_01.xlsx sheet=Missing cols=A:GRN No,B:Part No,C:Heat No rows=0 stale=2
TEST     fixture=grn_01.xlsx call=ListMissingCerts(GRN, Certs, Missing) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTA(Missing![GRN No]) = 4 AND ROWS_USED(Missing) = 4 AND COUNTIF(Missing![GRN No], "stale") = 0
TEST     fixture=grn_01.xlsx call=ListMissingCerts(GRN, Certs, Missing) assert=COUNTA(GRN![GRN No]) = PRE_COUNTA(GRN![GRN No]) AND COUNTA(GRN![Heat No]) = PRE_COUNTA(GRN![Heat No]) AND COUNTA(Certs![Heat No]) = PRE_COUNTA(Certs![Heat No]) AND VALUE(Missing![GRN No]2) = "G-1" AND VALUE(Missing![Part No]2) = "P-1" AND VALUE(Missing![Heat No]2) = "H-300" AND VALUE(Missing![GRN No]3) = "G-5" AND VALUE(Missing![GRN No]4) = "G-7" AND VALUE(Missing![Heat No]4) = " H-400 "
TEST     fixture=grn_01.xlsx call=ListMissingCerts(Certs, GRN, Missing) assert=RESULT = -1 AND ERROR_COUNT = 1
