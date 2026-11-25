VCL 0.7
TASK
# ASSUME: ITP Tracker headers in row 1 include "Inspection Status", whose values are Open or Closed
# ASSUME: a point shows when its Inspection Status is Open (compared with SAME); every other row is hidden, and a re-run shows reopened rows again
GOAL     Hide every closed point of the ITP Tracker so only the open points show, and return how many open points are visible.
PROC     ShowOpenInspectionPoints
KIND     Function
MODULE   modItp
ARGS     ws As Worksheet
RET      Long
VAR      cStatus As Long
VAR      shown As Long
PRE      COL($ws, "Inspection Status") > 0
MODIFIES $ws
OP       LET(cStatus, COL($ws, "Inspection Status"))
OP       FILTER(USED_RANGE($ws), col=cStatus, crit="Open")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cStatus) var=r
OP         IF SAME(CELL($ws, r, cStatus), "Open") THEN LET(shown, shown + 1)
OP       RETURN(shown)
PERF     screen_off
ERR      standard
FIXTURE  itp_01.xlsx sheet=Itp cols=A:ITP No,B:Point,C:Inspection Status rows=8 data=A:seq(1,1);B:text(10);C:values("Open"x5,"Closed"x3)+boundary("Open","Closed","Closed","Open","Open","Closed","Open","Open")
FIXTURE  itp_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=itp_01.xlsx call=ShowOpenInspectionPoints(Itp) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND VISIBLE_COUNT(Itp![Inspection Status]) = 6
TEST     fixture=itp_01.xlsx call=ShowOpenInspectionPoints(Itp) assert=COUNTA(Itp![ITP No]) = PRE_COUNTA(Itp![ITP No]) AND COUNTIF(Itp![Inspection Status], "Closed") = 3 AND VISIBLE_COUNT(Itp![ITP No]) = 6
TEST     fixture=itp_01.xlsx call=ShowOpenInspectionPoints(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
