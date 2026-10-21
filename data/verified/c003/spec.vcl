VCL 0.6
TASK
# ASSUME: NDE Report headers (row 1) include "Weld Seam No", "Welder ID" and "Result" (Accepted / Repair); one row per seam
# ASSUME: rows with a blank Welder ID or a blank Result are skipped; a blank Result means not inspected yet
# ASSUME: "Repair" is matched ignoring case and extra spaces
# ASSUME: repair rate = repaired seams / inspected seams per welder; Welder Performance is rebuilt every run, highest rate first
GOAL     Rebuild a Welder Performance sheet with each welder's inspected seams, repaired seams and repair rate, highest rate first.
PROC     BuildWelderRepairRate
KIND     Function
MODULE   modNde
ARGS     nde As Worksheet, out As Worksheet
RET      Long
VAR      total As Dictionary
VAR      repairs As Dictionary
VAR      cSeam As Long
VAR      cWelder As Long
VAR      cResult As Long
VAR      welderId As String
VAR      inspected As Long
VAR      repaired As Long
VAR      outRow As Long
PRE      COL($nde, "Weld Seam No") > 0; COL($nde, "Welder ID") > 0; COL($nde, "Result") > 0
MODIFIES $out
OP       LET(cSeam, COL($nde, "Weld Seam No"))
OP       LET(cWelder, COL($nde, "Welder ID"))
OP       LET(cResult, COL($nde, "Result"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Welder ID", "Seams Inspected", "Repairs", "Repair Rate")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($nde, cSeam) var=r
OP         IF NOT IS_EMPTY(CELL($nde, r, cWelder)) AND NOT IS_EMPTY(CELL($nde, r, cResult))
OP           LET(welderId, TRIM(CELL($nde, r, cWelder)))
OP           DICT_ADD(total, welderId, 1, mode=count)
OP           IF SAME(CELL($nde, r, cResult), "REPAIR") THEN DICT_ADD(repairs, welderId, 1, mode=count)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(total)
OP         DICT_GET(total, k, into=inspected)
OP         LET(repaired, 0)
OP         IF EXISTS(repairs, k) THEN DICT_GET(repairs, k, into=repaired)
OP         WRITE_ROW($out, outRow, k, inspected, repaired)
OP         SET_VALUE(CELL($out, outRow, 4), repaired / inspected)
OP         FORMAT(CELL($out, outRow, 4), numfmt="0.0%")
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=4, order=desc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  nde_repair_01.xlsx sheet=NDE cols=A:Weld Seam No,B:Welder ID,C:Result rows=60 data=A:seq(1,1);B:values("W-01"x20,"W-02"x19,"W-03"x15,"W-04"x1,""x5)+boundary("W-01","W-04");C:values("Accepted"x48,"Repair"x8,""x4)+boundary("Repair","Accepted")
FIXTURE  nde_repair_01.xlsx sheet=WelderPerf cols=A:Welder ID,B:Seams Inspected,C:Repairs,D:Repair Rate rows=0 stale=10
TEST     fixture=nde_repair_01.xlsx call=BuildWelderRepairRate(NDE, WelderPerf) runs=2 assert=RESULT = 4 AND COUNTA(WelderPerf![Welder ID]) = RESULT + 1 AND ERROR_COUNT = 0
TEST     fixture=nde_repair_01.xlsx call=BuildWelderRepairRate(NDE, WelderPerf) assert=SUM(WelderPerf![Repairs]) = PRE_COUNTIFS(NDE![Result], "Repair", NDE![Welder ID], "<>") AND SUM(WelderPerf![Seams Inspected]) = PRE_COUNTIFS(NDE![Result], "<>", NDE![Welder ID], "<>") - 1 AND COUNTIFS(WelderPerf![Welder ID], "W-04", WelderPerf![Repairs], 0) = 1 AND VALUE(WelderPerf![Repair Rate]2) >= VALUE(WelderPerf![Repair Rate]3)
TEST     fixture=nde_repair_01.xlsx call=BuildWelderRepairRate(NDE, WelderPerf) assert=COUNTA(NDE![Welder ID]) = PRE_COUNTA(NDE![Welder ID]) AND COUNTA(NDE![Result]) = PRE_COUNTA(NDE![Result])
TEST     fixture=nde_repair_01.xlsx call=BuildWelderRepairRate(WelderPerf, WelderPerf) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(WelderPerf![Welder ID]) = PRE_COUNTA(WelderPerf![Welder ID])
