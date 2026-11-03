VCL 0.6
TASK
# ASSUME: ITP Tracker headers in row 1 include "Inspection Status"; the real sheet is "ITP Tracker", the TESTs use Itp
# ASSUME: a status counts as Open or Closed when its text, ignoring case and spaces at both ends, is open or closed
# ASSUME: every other status, and every blank cell, is left exactly as it is
# ASSUME: a cell that already reads exactly Open or Closed is not changed and not counted
GOAL     Rewrite every Inspection Status that means Open or Closed, whatever its case or stray spaces, as exactly Open or Closed, leave other statuses alone and tell how many cells changed.
PROC     NormaliseInspectionStatus
KIND     Function
MODULE   modItp
ARGS     ws As Worksheet
RET      Long
VAR      cStatus As Long
VAR      changed As Long
PRE      COL($ws, "Inspection Status") > 0
MODIFIES $ws["Inspection Status"]
OP       LET(cStatus, COL($ws, "Inspection Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cStatus) var=r
OP         IF SAME(CELL($ws, r, cStatus), "Open")
OP           IF CELL($ws, r, cStatus) <> "Open"
OP             SET_VALUE(CELL($ws, r, cStatus), "Open")
OP             LET(changed, changed + 1)
OP         ELSE
OP           IF SAME(CELL($ws, r, cStatus), "Closed")
OP             IF CELL($ws, r, cStatus) <> "Closed"
OP               SET_VALUE(CELL($ws, r, cStatus), "Closed")
OP               LET(changed, changed + 1)
OP       MSG("Cells changed: " & changed)
OP       RETURN(changed)
ERR      standard
FIXTURE  itp_01.xlsx sheet=Itp cols=A:ITP No,B:Activity,C:Inspection Status,D:Inspector rows=15 data=A:seq(1,1);B:text(10);C:values("open"x1,"OPEN "x1," closed"x1,"Open"x1,"Closed"x1,"CLOSED"x1,"  Open  "x1,"Closed "x1,"Open - hold"x1,"Reopened"x1,""x1,"N/A"x1,"In Progress"x1,"closed"x1,"OPEN"x1)+boundary("open","OPEN "," closed","Open","Closed","CLOSED","  Open  ","Closed ","Open - hold","Reopened","","N/A","In Progress","closed","OPEN");D:values("QA1"x8,"QA2"x7)
FIXTURE  itp_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=itp_01.xlsx call=NormaliseInspectionStatus(Itp) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(Itp![Inspection Status], "Open") = 5 AND COUNTIF(Itp![Inspection Status], "Closed") = 5
TEST     fixture=itp_01.xlsx call=NormaliseInspectionStatus(Itp) assert=RESULT = 8 AND ERROR_COUNT = 0 AND COUNTA(Itp![Inspection Status]) = PRE_COUNTA(Itp![Inspection Status]) AND COUNTA(Itp![ITP No]) = PRE_COUNTA(Itp![ITP No]) AND COUNTA(Itp![Activity]) = PRE_COUNTA(Itp![Activity])
TEST     fixture=itp_01.xlsx call=NormaliseInspectionStatus(Itp) assert=VALUE(Itp![Inspection Status]2) = "Open" AND VALUE(Itp![Inspection Status]3) = "Open" AND VALUE(Itp![Inspection Status]4) = "Closed" AND VALUE(Itp![Inspection Status]5) = "Open" AND VALUE(Itp![Inspection Status]6) = "Closed" AND VALUE(Itp![Inspection Status]7) = "Closed" AND VALUE(Itp![Inspection Status]8) = "Open" AND VALUE(Itp![Inspection Status]9) = "Closed" AND VALUE(Itp![Inspection Status]15) = "Closed" AND VALUE(Itp![Inspection Status]16) = "Open"
TEST     fixture=itp_01.xlsx call=NormaliseInspectionStatus(Itp) assert=VALUE(Itp![Inspection Status]10) = "Open - hold" AND VALUE(Itp![Inspection Status]11) = "Reopened" AND VALUE(Itp![Inspection Status]12) = "" AND VALUE(Itp![Inspection Status]13) = "N/A" AND VALUE(Itp![Inspection Status]14) = "In Progress"
TEST     fixture=itp_01.xlsx call=NormaliseInspectionStatus(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
