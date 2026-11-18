VCL 0.6
TASK
# ASSUME: NCR Register headers in row 1 include NCR No, Department and Status
# ASSUME: a status counts as open when it is Open ignoring case and spaces, and as closed when it is Closed or Close ignoring case and spaces; nothing else is matched
# ASSUME: a cell that already reads exactly Open or Closed is not changed and not counted
GOAL     Rewrite every open-like status in the NCR Register as exactly Open and every closed-like status as exactly Closed, and return how many cells changed.
PROC     NormalizeNcrStatus
KIND     Function
MODULE   modNcr
ARGS     ncr As Worksheet
RET      Long
VAR      cStatus As Long
VAR      changed As Long
PRE      COL($ncr, "Status") > 0
MODIFIES $ncr["Status"]
OP       LET(cStatus, COL($ncr, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, cStatus) var=r
OP         SKIP_IF IS_EMPTY(CELL($ncr, r, cStatus))
OP         IF SAME(CELL($ncr, r, cStatus), "Open")
OP           IF CELL($ncr, r, cStatus) <> "Open"
OP             SET_VALUE(CELL($ncr, r, cStatus), "Open")
OP             LET(changed, changed + 1)
OP         ELSE
OP           IF SAME(CELL($ncr, r, cStatus), "Closed") OR SAME(CELL($ncr, r, cStatus), "Close")
OP             IF CELL($ncr, r, cStatus) <> "Closed"
OP               SET_VALUE(CELL($ncr, r, cStatus), "Closed")
OP               LET(changed, changed + 1)
OP       MSG("Cells changed: " & changed)
OP       RETURN(changed)
PERF     screen_off
ERR      standard
FIXTURE  ncr_01.xlsx sheet=NcrReg cols=A:NCR No,B:Department,C:Status rows=11 data=A:seq(1,1);B:values("QA"x6,"Weld"x5);C:values("Open"x1,"x"x10)+boundary("Open","open ","",""" OPEN""","Closed","CLOSED","close"," Close ","closed ","Pending","Reopen")
FIXTURE  ncr_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=ncr_01.xlsx call=NormalizeNcrStatus(NcrReg) assert=RESULT = 6 AND ERROR_COUNT = 0 AND VALUE(NcrReg![Status]3) = "Open" AND VALUE(NcrReg![Status]5) = "Open" AND VALUE(NcrReg![Status]7) = "Closed" AND VALUE(NcrReg![Status]8) = "Closed" AND VALUE(NcrReg![Status]9) = "Closed" AND VALUE(NcrReg![Status]10) = "Closed"
TEST     fixture=ncr_01.xlsx call=NormalizeNcrStatus(NcrReg) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE(NcrReg![Status]2) = "Open" AND VALUE(NcrReg![Status]6) = "Closed" AND VALUE(NcrReg![Status]11) = "Pending" AND VALUE(NcrReg![Status]12) = "Reopen" AND COUNTA(NcrReg![Status]) = PRE_COUNTA(NcrReg![Status]) AND COUNTA(NcrReg![Department]) = PRE_COUNTA(NcrReg![Department]) AND COUNTA(NcrReg![NCR No]) = PRE_COUNTA(NcrReg![NCR No])
TEST     fixture=ncr_01.xlsx call=NormalizeNcrStatus(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
