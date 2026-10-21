VCL 0.6
TASK
# ASSUME: the real sheet is "ITP Tracker"; its row-1 headers include "Inspection Status" (the TESTs call the sheet Itp)
# ASSUME: Inspection Status holds only Open or Closed, matched ignoring case and extra spaces
# ASSUME: FILTER keeps only the rows equal to one value, so a row with a blank status or any status but Open is hidden too
# ASSUME: hiding is not deleting; every run shows or hides every row again from its current status, so a reopened row shows again
GOAL     Hide every closed row of the ITP Tracker so that only open inspection points show, and return how many rows are visible below the headings.
PROC     ShowOpenInspectionPoints
KIND     Function
MODULE   modItp
ARGS     ws As Worksheet
RET      Long
VAR      cStatus As Long
VAR      visible As Long
PRE      COL($ws, "Inspection Status") > 0
MODIFIES $ws
OP       LET(cStatus, COL($ws, "Inspection Status"))
OP       FILTER(USED_RANGE($ws), col=cStatus, crit="Open")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cStatus) var=r
OP         IF SAME(CELL($ws, r, cStatus), "Open")
OP           LET(visible, visible + 1)
OP       MSG("Visible inspection points: " & visible)
OP       RETURN(visible)
ERR      standard
FIXTURE  itp_01.xlsx sheet=Itp cols=A:ITP No,B:Point,C:Inspection Status rows=10 data=A:seq(1,1);B:text(10);C:values("Open"x4," open "x1,"Closed"x4,"CLOSED"x1)+boundary(" open ","CLOSED","Open","Closed")
FIXTURE  itp_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=itp_01.xlsx call=ShowOpenInspectionPoints(Itp) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND VISIBLE_COUNT(Itp![Inspection Status]) = 6 AND VISIBLE_COUNT(Itp![ITP No]) = 6
TEST     fixture=itp_01.xlsx call=ShowOpenInspectionPoints(Itp) assert=RESULT = 5 AND COUNTA(Itp![ITP No]) = PRE_COUNTA(Itp![ITP No]) AND COUNTIF(Itp![Inspection Status], "Closed") = PRE_COUNTIF(Itp![Inspection Status], "Closed") AND VISIBLE_COUNT(Itp![Point]) = 6
TEST     fixture=itp_01.xlsx call=ShowOpenInspectionPoints(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1 AND VISIBLE_COUNT(Notes![Note]) = PRE_VISIBLE_COUNT(Notes![Note])
