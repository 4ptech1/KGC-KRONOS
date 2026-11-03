VCL 0.6
TASK
# ASSUME: Welder Quals headers in row 1 include "Process" and "Expiry Date"; the real sheets are "Welder Quals" and "Qual Summary", the TESTs use Quals and QualSummary
# ASSUME: the Qual Summary sheet already exists and is rebuilt on every run with the headers Process, Valid Over 60 Days, Expiring Within 60 Days and Expired
# ASSUME: valid for more than 60 days means expiry after today plus 60; expiring within the next 60 days means today up to and including today plus 60; expired means before today
# ASSUME: processes are matched ignoring case and stray spaces and shown as first typed; a row with no process or an Expiry Date that is not a real date is ignored
GOAL     Rebuild the Qual Summary sheet with one row per welding process, in alphabetical order, counting qualifications valid for over 60 days, expiring within 60 days and already expired.
PROC     BuildQualSummary
KIND     Function
MODULE   modQuals
ARGS     quals As Worksheet, out As Worksheet
RET      Long
VAR      cProc As Long
VAR      cExp As Long
VAR      p As String
VAR      valid As Dictionary
VAR      soon As Dictionary
VAR      gone As Dictionary
VAR      nValid As Long
VAR      nSoon As Long
VAR      nGone As Long
VAR      outRow As Long
PRE      COL($quals, "Process") > 0; COL($quals, "Expiry Date") > 0
MODIFIES $out
OP       LET(cProc, COL($quals, "Process"))
OP       LET(cExp, COL($quals, "Expiry Date"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($quals, cProc) var=r
OP         SKIP_IF IS_EMPTY(CELL($quals, r, cProc))
OP         SKIP_IF NOT IS_DATE(CELL($quals, r, cExp))
OP         LET(p, TRIM(CELL($quals, r, cProc)))
OP         DICT_ADD(valid, p, 0, mode=sum)
OP         DICT_ADD(soon, p, 0, mode=sum)
OP         DICT_ADD(gone, p, 0, mode=sum)
OP         IF CELL($quals, r, cExp) < TODAY
OP           DICT_ADD(gone, p, 1, mode=sum)
OP         ELSE
OP           IF CELL($quals, r, cExp) <= TODAY + 60
OP             DICT_ADD(soon, p, 1, mode=sum)
OP           ELSE
OP             DICT_ADD(valid, p, 1, mode=sum)
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Process", "Valid Over 60 Days", "Expiring Within 60 Days", "Expired")
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(valid)
OP         DICT_GET(valid, k, into=nValid)
OP         DICT_GET(soon, k, into=nSoon)
OP         DICT_GET(gone, k, into=nGone)
OP         WRITE_ROW($out, outRow, k, nValid, nSoon, nGone)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       RETURN(outRow - 2)
ERR      standard
FIXTURE  quals_01.xlsx sheet=Quals cols=A:Welder ID,B:Name,C:Process,D:Expiry Date rows=13 data=A:seq(1,1);B:text(8);C:values("SMAW"x3,"smaw "x1,"GTAW"x2,"gtaw"x1,"FCAW"x3,""x1,"GMAW"x1,"SAW"x1)+boundary("SMAW","SMAW","smaw ","SMAW","GTAW","GTAW","gtaw","FCAW","FCAW","","GMAW","SAW","FCAW");D:reldates(-200,400)+boundary(-1,0,60,61,59,400,-30,120,-200,30,"",1,"")
FIXTURE  quals_01.xlsx sheet=QualSummary cols=A:Process,B:Valid Over 60 Days,C:Expiring Within 60 Days,D:Expired rows=0 stale=6
FIXTURE  quals_02.xlsx sheet=Quals cols=A:Welder ID,B:Name,C:Process,D:Expiry Date rows=4 data=A:seq(1,1);B:text(8);C:values("SMAW"x2,"GTAW"x2);D:text(10)
FIXTURE  quals_02.xlsx sheet=QualSummary cols=A:Process,B:Valid Over 60 Days,C:Expiring Within 60 Days,D:Expired rows=0 stale=2
TEST     fixture=quals_01.xlsx call=BuildQualSummary(Quals, QualSummary) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND ROWS_USED(QualSummary) = 5 AND COUNTA(QualSummary![Process]) = 5 AND COUNTIF(QualSummary![Process], "stale") = 0 AND VALUE(QualSummary![Process]2) = "FCAW" AND VALUE(QualSummary![Process]3) = "GTAW" AND VALUE(QualSummary![Process]4) = "SAW" AND VALUE(QualSummary![Process]5) = "SMAW"
TEST     fixture=quals_01.xlsx call=BuildQualSummary(Quals, QualSummary) assert=VALUE(QualSummary![Valid Over 60 Days]2) = 1 AND VALUE(QualSummary![Expiring Within 60 Days]2) = 0 AND VALUE(QualSummary![Expired]2) = 1 AND VALUE(QualSummary![Valid Over 60 Days]3) = 1 AND VALUE(QualSummary![Expiring Within 60 Days]3) = 1 AND VALUE(QualSummary![Expired]3) = 1 AND VALUE(QualSummary![Valid Over 60 Days]4) = 0 AND VALUE(QualSummary![Expiring Within 60 Days]4) = 1 AND VALUE(QualSummary![Expired]4) = 0 AND VALUE(QualSummary![Valid Over 60 Days]5) = 1 AND VALUE(QualSummary![Expiring Within 60 Days]5) = 2 AND VALUE(QualSummary![Expired]5) = 1
TEST     fixture=quals_01.xlsx call=BuildQualSummary(Quals, QualSummary) assert=RESULT = 4 AND COUNTA(Quals![Welder ID]) = PRE_COUNTA(Quals![Welder ID]) AND COUNTA(Quals![Process]) = PRE_COUNTA(Quals![Process]) AND COUNTA(Quals![Expiry Date]) = PRE_COUNTA(Quals![Expiry Date])
TEST     fixture=quals_01.xlsx call=BuildQualSummary(QualSummary, Quals) assert=RESULT = -1 AND ERROR_COUNT = 1
TEST     fixture=quals_02.xlsx call=BuildQualSummary(Quals, QualSummary) assert=RESULT = 0 AND ERROR_COUNT = 0 AND ROWS_USED(QualSummary) = 1 AND COUNTA(QualSummary![Process]) = 1 AND COUNTIF(QualSummary![Process], "stale") = 0
