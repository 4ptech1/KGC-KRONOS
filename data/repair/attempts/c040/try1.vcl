VCL 0.6
TASK
# ASSUME: WeldLog headers in row 1 include "Weld Seam No" and "Visual"; NdeReport headers include "Weld Seam No" and "Result"
# ASSUME: the real sheets are "Weld Log", "NDE Report" and "Not Inspected"; the Not Inspected sheet exists and is rebuilt on every run
# ASSUME: a seam has an NDE row when its trimmed Weld Seam No appears in the NDE Report, ignoring case; the Result is not looked at
# ASSUME: Visual is matched ignoring case and extra spaces; rows with a blank Weld Seam No are ignored
# ASSUME: the listed seams are written trimmed, in weld log order, one row for each weld log row; the sheet has the heading "Weld Seam No"
GOAL     List on the Not Inspected sheet every weld log seam that has Visual Accepted but no row in the NDE Report.
PROC     ListSeamsWithoutNde
KIND     Function
MODULE   modNdeCheck
ARGS     weld As Worksheet, nde As Worksheet, out As Worksheet
RET      Long
VAR      cWSeam As Long
VAR      cVisual As Long
VAR      cNSeam As Long
VAR      seen As Dictionary
VAR      outRow As Long
PRE      COL($weld, "Weld Seam No") > 0; COL($weld, "Visual") > 0; COL($nde, "Weld Seam No") > 0
MODIFIES $out
OP       LET(cWSeam, COL($weld, "Weld Seam No"))
OP       LET(cVisual, COL($weld, "Visual"))
OP       LET(cNSeam, COL($nde, "Weld Seam No"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($nde, cNSeam) var=n
OP         SKIP_IF IS_EMPTY(CELL($nde, n, cNSeam))
OP         DICT_ADD(seen, TRIM(CELL($nde, n, cNSeam)), 1, mode=keep)
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Weld Seam No")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($weld, cWSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($weld, r, cWSeam))
OP         IF SAME(CELL($weld, r, cVisual), "Accepted") AND NOT EXISTS(seen, TRIM(CELL($weld, r, cWSeam)))
OP           WRITE_ROW($out, outRow, TRIM(CELL($weld, r, cWSeam)))
OP           LET(outRow, outRow + 1)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  nde_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Visual rows=12 data=A:values("W-001"x1,"W-002"x1,"W-003"x1," W-004 "x1,"W-005"x1,""x1,"W-007"x1,"w-008"x1," W-009"x1,"W-010"x1,"W-011"x1,"W-012"x1)+boundary("W-001","W-002","W-003"," W-004 ","W-005","","W-007","w-008"," W-009","W-010","W-011","W-012");B:values("Accepted"x8,"Rejected"x2," accepted "x1,""x1)+boundary("Accepted","Accepted","Rejected","Accepted"," accepted ","Accepted","","Accepted","Accepted","Accepted","Rejected","Accepted")
FIXTURE  nde_01.xlsx sheet=NdeReport cols=A:Weld Seam No,B:Result rows=7 data=A:values("W-001"x2,"W-004"x1,"W-008"x1," W-010 "x1,""x1,"W-099"x1)+boundary("W-001","W-004","W-008"," W-010 ","","W-099","W-001");B:values("Pass"x5,"Fail"x2)
FIXTURE  nde_01.xlsx sheet=NotInspected cols=A:Weld Seam No rows=0 stale=6
TEST     fixture=nde_01.xlsx call=ListSeamsWithoutNde(WeldLog, NdeReport, NotInspected) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND ROWS_USED(NotInspected) = 5 AND COUNTIF(NotInspected![Weld Seam No], "stale") = 0 AND VALUE(NotInspected![Weld Seam No]2) = "W-002" AND VALUE(NotInspected![Weld Seam No]3) = "W-005" AND VALUE(NotInspected![Weld Seam No]4) = "W-009" AND VALUE(NotInspected![Weld Seam No]5) = "W-012"
TEST     fixture=nde_01.xlsx call=ListSeamsWithoutNde(WeldLog, NdeReport, NotInspected) assert=COUNTA(NotInspected![Weld Seam No]) = 5 AND COUNTA(WeldLog![Weld Seam No]) = PRE_COUNTA(WeldLog![Weld Seam No]) AND COUNTA(WeldLog![Visual]) = PRE_COUNTA(WeldLog![Visual]) AND COUNTA(NdeReport![Weld Seam No]) = PRE_COUNTA(NdeReport![Weld Seam No]) AND COUNTA(NdeReport![Result]) = PRE_COUNTA(NdeReport![Result]) AND ERROR_COUNT = 0
TEST     fixture=nde_01.xlsx call=ListSeamsWithoutNde(NdeReport, WeldLog, NotInspected) assert=RESULT = -1 AND ERROR_COUNT = 1
