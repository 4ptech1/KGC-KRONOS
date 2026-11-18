VCL 0.6
TASK
# ASSUME: Welder Quals headers in row 1 include "Welder ID", "Process" and "Expiry Date"
# ASSUME: the real sheets are "Welder Quals" and "Expiring Quals"; Expiring Quals exists and is rebuilt on every run
# ASSUME: Expiry Date holds real dates; a row with a blank or non-date Expiry Date is skipped
# ASSUME: expires in the next 60 days includes exactly 60 days from today; Days Left is the whole days from today, 0 for today
GOAL     Rebuild the Expiring Quals sheet with every qualification that has expired or expires within 60 days, with Days Left, soonest first.
PROC     ListExpiringQuals
KIND     Function
MODULE   modQuals
ARGS     quals As Worksheet, out As Worksheet
RET      Long
VAR      cId As Long
VAR      cProc As Long
VAR      cExp As Long
VAR      outRow As Long
PRE      COL($quals, "Welder ID") > 0; COL($quals, "Process") > 0; COL($quals, "Expiry Date") > 0
MODIFIES $out
OP       LET(cId, COL($quals, "Welder ID"))
OP       LET(cProc, COL($quals, "Process"))
OP       LET(cExp, COL($quals, "Expiry Date"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Welder ID", "Process", "Expiry Date", "Days Left")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($quals, cExp) var=r
OP         SKIP_IF NOT IS_DATE(CELL($quals, r, cExp))
OP         IF CELL($quals, r, cExp) <= TODAY + 60
OP           WRITE_ROW($out, outRow, CELL($quals, r, cId), CELL($quals, r, cProc), CELL($quals, r, cExp), CELL($quals, r, cExp) - TODAY)
OP           LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=4, order=asc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off
ERR      standard
FIXTURE  quals_01.xlsx sheet=Quals cols=A:Welder ID,B:Process,C:Expiry Date rows=9 data=A:seq(1,1);B:values("GTAW"x5,"SMAW"x4);C:reldates(-90,200)+boundary(60,61,-1,0,1,"",-90,59,200)
FIXTURE  quals_01.xlsx sheet=Expiring cols=A:Welder ID,B:Process,C:Expiry Date,D:Days Left rows=0 stale=2
TEST     fixture=quals_01.xlsx call=ListExpiringQuals(Quals, Expiring) runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND COUNTA(Expiring![Welder ID]) = 7 AND COUNTIF(Expiring![Welder ID], "stale") = 0 AND VALUE(Expiring![Days Left]2) = -90 AND VALUE(Expiring![Days Left]3) = -1 AND VALUE(Expiring![Days Left]4) = 0 AND VALUE(Expiring![Days Left]5) = 1 AND VALUE(Expiring![Days Left]6) = 59 AND VALUE(Expiring![Days Left]7) = 60
TEST     fixture=quals_01.xlsx call=ListExpiringQuals(Quals, Expiring) assert=COUNTA(Quals![Welder ID]) = PRE_COUNTA(Quals![Welder ID]) AND COUNTA(Quals![Expiry Date]) = PRE_COUNTA(Quals![Expiry Date])
TEST     fixture=quals_01.xlsx call=ListExpiringQuals(Expiring, Quals) assert=RESULT = -1 AND ERROR_COUNT = 1
