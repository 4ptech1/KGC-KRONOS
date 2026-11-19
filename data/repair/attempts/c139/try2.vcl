VCL 0.6
TASK
# ASSUME: Weld Log headers in row 1 include "Weld Seam No", "Welder ID" and "Process"; Welder Quals has "Welder ID" and "Process"
# ASSUME: a welder is qualified when Welder Quals has a row with the same Welder ID and Process, ignoring case and stray spaces
# ASSUME: the Expiry Date is not checked, and a Welder Quals row with no Welder ID or no Process is ignored
# ASSUME: a weld with no Welder ID or no Process has no qual; a Weld Log row with no Weld Seam No is skipped and left as it is
# ASSUME: Qual Check is added at the end of row 1 when missing; a qualified weld gets an empty Qual Check
GOAL     Mark each weld in the Weld Log with No qual in Qual Check when its welder has no qualification for that process, clear the others, and return how many have none.
PROC     CheckWeldQuals
KIND     Function
MODULE   modWeldQuals
ARGS     log As Worksheet, quals As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cWelder As Long
VAR      cProc As Long
VAR      cCheck As Long
VAR      qWelder As Long
VAR      qProc As Long
VAR      held As Dictionary
VAR      n As Long
PRE      COL($log, "Weld Seam No") > 0; COL($log, "Welder ID") > 0; COL($log, "Process") > 0; COL($quals, "Welder ID") > 0; COL($quals, "Process") > 0
MODIFIES $log
OP       LET(cSeam, COL($log, "Weld Seam No"))
OP       LET(cWelder, COL($log, "Welder ID"))
OP       LET(cProc, COL($log, "Process"))
OP       LET(qWelder, COL($quals, "Welder ID"))
OP       LET(qProc, COL($quals, "Process"))
OP       LET(cCheck, COL($log, "Qual Check"))
OP       IF cCheck = 0
OP         LET(cCheck, LASTCOL($log, 1) + 1)
OP         SET_VALUE(CELL($log, 1, cCheck), "Qual Check")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($quals, qWelder) var=q
OP         SKIP_IF IS_EMPTY(CELL($quals, q, qWelder)) OR IS_EMPTY(CELL($quals, q, qProc))
OP         DICT_ADD(held, TRIM(CELL($quals, q, qWelder)) & "|" & TRIM(CELL($quals, q, qProc)), 1, mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($log, cSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($log, r, cSeam))
OP         IF EXISTS(held, TRIM(CELL($log, r, cWelder)) & "|" & TRIM(CELL($log, r, cProc)))
OP           SET_VALUE(CELL($log, r, cCheck), "")
OP         ELSE
OP           SET_VALUE(CELL($log, r, cCheck), "No qual")
OP           LET(n, n + 1)
OP       MSG("Welds with no qual: " & n)
OP       RETURN(n)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  weldqual_01.xlsx sheet=Log cols=A:Weld Seam No,B:Welder ID,C:Process,D:Qual Check rows=8 data=A:values("S1"x1,"S2"x1,"S3"x1,"S4"x1,"S5"x1,"S6"x1,""x1,"S8"x1)+boundary("S1","S2","S3","S4","S5","S6","","S8");B:values("W01"x2," w01"x1,"W02"x1,"W04"x2,"W03"x1,""x1)+boundary("W01"," w01","W02","W04","W03","","W04","W01");C:values("GTAW"x4,"smaw"x1,"SMAW"x3)+boundary("GTAW","smaw","GTAW","SMAW","GTAW","GTAW","SMAW","SMAW");D:values("No qual"x2,"stale text"x1,"keep"x1,""x4)+boundary("No qual","stale text","","","No qual","","keep","")
FIXTURE  weldqual_01.xlsx sheet=Quals cols=A:Welder ID,B:Process,C:Expiry Date rows=5 data=A:values("W01"x2,"w02 "x1,""x1," W03 "x1)+boundary("W01","w02 ","W01",""," W03 ");B:values("GTAW"x2,"SMAW"x2," gtaw "x1)+boundary("GTAW","SMAW","SMAW","GTAW"," gtaw ");C:dates(2026-01-01,2027-12-31)
FIXTURE  weldqual_02.xlsx sheet=Log cols=A:Weld Seam No,B:Welder ID,C:Process rows=3 data=A:seq(1,1);B:values("W01"x2,"W09"x1)+boundary("W01","W09","W01");C:values("GTAW"x2,"SMAW"x1)+boundary("GTAW","GTAW","SMAW")
FIXTURE  weldqual_02.xlsx sheet=Quals cols=A:Welder ID,B:Process,C:Expiry Date rows=2 data=A:values("W01"x2);B:values("GTAW"x1,"SMAW"x1)+boundary("GTAW","SMAW");C:dates(2026-01-01,2027-12-31)
TEST     fixture=weldqual_01.xlsx call=CheckWeldQuals(Log, Quals) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Log![Qual Check], "No qual") = 3 AND COUNTIF(Log![Qual Check], "stale text") = 0 AND VALUE(Log![Qual Check]2) = "" AND VALUE(Log![Qual Check]3) = "" AND VALUE(Log![Qual Check]4) = "No qual" AND VALUE(Log![Qual Check]5) = "No qual" AND VALUE(Log![Qual Check]6) = "" AND VALUE(Log![Qual Check]7) = "No qual" AND VALUE(Log![Qual Check]8) = "keep" AND VALUE(Log![Qual Check]9) = ""
TEST     fixture=weldqual_01.xlsx call=CheckWeldQuals(Log, Quals) assert=ERROR_COUNT = 0 AND COUNTA(Log![Weld Seam No]) = PRE_COUNTA(Log![Weld Seam No]) AND COUNTA(Log![Welder ID]) = PRE_COUNTA(Log![Welder ID]) AND COUNTA(Quals![Welder ID]) = PRE_COUNTA(Quals![Welder ID]) AND COUNTA(Quals![Expiry Date]) = PRE_COUNTA(Quals![Expiry Date])
TEST     fixture=weldqual_02.xlsx call=CheckWeldQuals(Log, Quals) runs=2 assert=RESULT = 1 AND ERROR_COUNT = 0 AND VALUE(Log!D1) = "Qual Check" AND COUNTA(Log![Qual Check]) = 2 AND VALUE(Log![Qual Check]3) = "No qual" AND COUNTA(Log![Weld Seam No]) = PRE_COUNTA(Log![Weld Seam No])
TEST     fixture=weldqual_01.xlsx call=CheckWeldQuals(Quals, Log) assert=RESULT = -1 AND ERROR_COUNT = 1
