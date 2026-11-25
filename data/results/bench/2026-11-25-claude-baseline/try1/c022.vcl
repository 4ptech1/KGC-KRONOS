VCL 0.7
TASK
# ASSUME: Weld Log headers in row 1 include "Welder ID", "Process" and "Remark"; Welder Quals headers include "Welder ID" and "Process"
# ASSUME: a weld is qualified when some Welder Quals row has the same Welder ID and Process (compared with SAME); expiry dates are not checked
# ASSUME: Remark is set to "Not qualified" or cleared on every weld row
GOAL     Mark every weld whose welder has no qualification row for its process as Not qualified in Remark, clear Remark on the others and return how many are Not qualified.
PROC     FlagUnqualifiedWelds
KIND     Function
MODULE   modQuals
ARGS     log As Worksheet, quals As Worksheet
RET      Long
VAR      cWelder As Long
VAR      cProc As Long
VAR      cRemark As Long
VAR      qWelder As Long
VAR      qProc As Long
VAR      found As Boolean
VAR      flagged As Long
PRE      COL($log, "Welder ID") > 0; COL($log, "Process") > 0; COL($log, "Remark") > 0; COL($quals, "Welder ID") > 0; COL($quals, "Process") > 0
MODIFIES $log["Remark"]
OP       LET(cWelder, COL($log, "Welder ID"))
OP       LET(cProc, COL($log, "Process"))
OP       LET(cRemark, COL($log, "Remark"))
OP       LET(qWelder, COL($quals, "Welder ID"))
OP       LET(qProc, COL($quals, "Process"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($log, cWelder) var=r
OP         LET(found, FALSE)
OP         LOOP_ROWS dir=down from=2 to=LASTROW($quals, qWelder) var=q
OP           IF SAME(CELL($log, r, cWelder), CELL($quals, q, qWelder)) AND SAME(CELL($log, r, cProc), CELL($quals, q, qProc)) THEN LET(found, TRUE)
OP         IF found
OP           SET_VALUE(CELL($log, r, cRemark), "")
OP         ELSE
OP           SET_VALUE(CELL($log, r, cRemark), "Not qualified")
OP           LET(flagged, flagged + 1)
OP       RETURN(flagged)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  weld_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Welder ID,C:Process,D:Remark rows=6 data=A:seq(1,1);B:values("W1"x3,"W2"x2,"W3"x1)+boundary("W1","W1","W2","W2","W3","W1");C:values("GTAW"x4,"SMAW"x2)+boundary("GTAW","SMAW","GTAW","SMAW","GTAW","GTAW");D:values("old"x6)
FIXTURE  weld_01.xlsx sheet=Quals cols=A:Welder ID,B:Process,C:Expiry Date rows=3 data=A:values("W1"x2,"W2"x1)+boundary("W1","W1","W2");B:values("GTAW"x2,"SMAW"x1)+boundary("GTAW","SMAW","GTAW");C:dates(2026-01-01,2026-12-31)
TEST     fixture=weld_01.xlsx call=FlagUnqualifiedWelds(WeldLog, Quals) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND COUNTIF(WeldLog![Remark], "Not qualified") = 2 AND COUNTIF(WeldLog![Remark], "old") = 0
TEST     fixture=weld_01.xlsx call=FlagUnqualifiedWelds(WeldLog, Quals) assert=COUNTA(WeldLog![Welder ID]) = PRE_COUNTA(WeldLog![Welder ID]) AND COUNTA(Quals![Welder ID]) = PRE_COUNTA(Quals![Welder ID]) AND VALUE(WeldLog![Remark]2) = "" AND VALUE(WeldLog![Remark]5) = "Not qualified"
TEST     fixture=weld_01.xlsx call=FlagUnqualifiedWelds(Quals, WeldLog) assert=RESULT = -1 AND ERROR_COUNT = 1
