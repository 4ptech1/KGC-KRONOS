VCL 0.6
TASK
# ASSUME: Weld Log headers (row 1) include "Weld Seam No", "Welder ID", "Process" and "Remark"; Welder Quals includes "Welder ID" and "Process"
# ASSUME: a welder is qualified when Welder Quals has a row with the same Welder ID and Process; Expiry Date is not looked at, so an expired row still counts
# ASSUME: Welder ID and Process are matched after trimming both ends and ignoring case
# ASSUME: a weld with a blank Welder ID or Process has no qualification row, so it is Not qualified; a row with a blank Weld Seam No is not a weld and is left alone
# ASSUME: a Welder Quals row with a blank Welder ID or Process counts for nobody
# ASSUME: whatever a qualified weld's Remark said before is cleared
GOAL     Write Not qualified in the Weld Log Remark of every weld whose welder has no qualification for its process, clear the Remark of the others and return how many welds are Not qualified.
PROC     FlagUnqualifiedWelds
KIND     Function
MODULE   modQuals
ARGS     log As Worksheet, quals As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cWelder As Long
VAR      cProc As Long
VAR      cRemark As Long
VAR      cQWelder As Long
VAR      cQProc As Long
VAR      held As Dictionary
VAR      n As Long
PRE      COL($log, "Weld Seam No") > 0; COL($log, "Welder ID") > 0; COL($log, "Process") > 0; COL($log, "Remark") > 0; COL($quals, "Welder ID") > 0; COL($quals, "Process") > 0
MODIFIES $log["Remark"]
OP       LET(cSeam, COL($log, "Weld Seam No"))
OP       LET(cWelder, COL($log, "Welder ID"))
OP       LET(cProc, COL($log, "Process"))
OP       LET(cRemark, COL($log, "Remark"))
OP       LET(cQWelder, COL($quals, "Welder ID"))
OP       LET(cQProc, COL($quals, "Process"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($quals, cQWelder) var=r
OP         SKIP_IF IS_EMPTY(CELL($quals, r, cQWelder)) OR IS_EMPTY(CELL($quals, r, cQProc))
OP         DICT_ADD(held, TRIM(CELL($quals, r, cQWelder)) & "|" & TRIM(CELL($quals, r, cQProc)), 1, mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($log, cSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($log, r, cSeam))
OP         IF EXISTS(held, TRIM(CELL($log, r, cWelder)) & "|" & TRIM(CELL($log, r, cProc)))
OP           CLEAR(CELL($log, r, cRemark), what=contents)
OP         ELSE
OP           SET_VALUE(CELL($log, r, cRemark), "Not qualified")
OP           LET(n, n + 1)
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  welds_01.xlsx sheet=Log cols=A:Weld Seam No,B:Welder ID,C:Process,D:Remark rows=10 data=A:values("S-1"x1,"S-2"x1,"S-3"x1,"S-4"x1,""x1,"S-6"x1,"S-7"x1,"S-8"x1,"S-9"x1,"S-10"x1)+boundary("S-1","S-2","S-3","S-4","","S-6","S-7","S-8","S-9","S-10");B:values("W-1"x2," w-1 "x1,"W-2"x2,"W-9"x2,""x1,"W-3"x1,"W-5"x1)+boundary("W-1","W-1"," w-1 ","W-2","W-9","W-9","W-2","","W-3","W-5");C:values("GTAW"x3," gtaw "x1,"SMAW"x4,"FCAW"x1,""x1)+boundary("SMAW","GTAW"," gtaw ","GTAW","FCAW","SMAW","SMAW","GTAW","","SMAW");D:values(""x5,"Not qualified"x2,"Re-test"x1,"keep"x1,"old"x1)+boundary("","","Not qualified","Re-test","keep","","Not qualified","","old","")
FIXTURE  welds_01.xlsx sheet=Quals cols=A:Welder ID,B:Process,C:Expiry Date rows=6 data=A:values("W-1"x1,"W-2"x1,""x1,"W-3"x1," W-1 "x1," W-5 "x1)+boundary("W-2","W-1","","W-3"," W-1 "," W-5 ");B:values("GTAW"x2,"SMAW"x1,""x1,"gtaw "x1,"SMAW "x1)+boundary("SMAW","GTAW","GTAW","","gtaw ","SMAW ");C:dates(2025-01-01,2027-12-31)
TEST     fixture=welds_01.xlsx call=FlagUnqualifiedWelds(Log, Quals) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTIF(Log![Remark], "Not qualified") = 5 AND COUNTA(Log![Remark]) = 7 AND COUNTIF(Log![Remark], "Re-test") = 0 AND COUNTIF(Log![Remark], "old") = 0
TEST     fixture=welds_01.xlsx call=FlagUnqualifiedWelds(Log, Quals) assert=COUNTA(Log![Weld Seam No]) = PRE_COUNTA(Log![Weld Seam No]) AND COUNTA(Log![Welder ID]) = PRE_COUNTA(Log![Welder ID]) AND COUNTA(Quals![Welder ID]) = PRE_COUNTA(Quals![Welder ID]) AND COUNTIF(Log![Remark], "keep") = 1 AND COUNTIFS(Log![Weld Seam No], "S-3", Log![Remark], "<>") = 0 AND COUNTIFS(Log![Weld Seam No], "S-7", Log![Remark], "<>") = 0 AND COUNTIFS(Log![Weld Seam No], "S-1", Log![Remark], "Not qualified") = 1 AND COUNTIFS(Log![Weld Seam No], "S-4", Log![Remark], "Not qualified") = 1 AND COUNTIFS(Log![Weld Seam No], "S-8", Log![Remark], "Not qualified") = 1
TEST     fixture=welds_01.xlsx call=FlagUnqualifiedWelds(Quals, Log) assert=RESULT = -1 AND ERROR_COUNT = 1
