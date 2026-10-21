VCL 0.6
TASK
# ASSUME: NCR Register headers in row 1 include "NCR No" and "Status"
# ASSUME: NCR-104 is matched as a whole value, ignoring case and extra spaces; the first match wins
# ASSUME: NCR-104 appears only in the NCR No column, so the first match anywhere on the sheet is its row
# ASSUME: Status is overwritten with the text Reviewed whatever it held before
# ASSUME: the function returns the sheet row number of the NCR, or 0 when it is not there
GOAL     Find the row of NCR-104 on the NCR Register, write Reviewed in its Status column and return that row number, or 0 if it is not there.
PROC     MarkNcr104Reviewed
KIND     Function
MODULE   modNcr
ARGS     ws As Worksheet
RET      Long
VAR      cStatus As Long
VAR      hit As Range
VAR      hitRow As Long
PRE      COL($ws, "NCR No") > 0; COL($ws, "Status") > 0
MODIFIES $ws["Status"]
OP       LET(cStatus, COL($ws, "Status"))
OP       LET(hit, FIND(USED_RANGE($ws), "NCR-104", match=whole))
OP       IF IS_NOTHING(hit)
OP         RETURN(0)
OP       LET(hitRow, ROW_OF(hit))
OP       SET_VALUE(CELL($ws, hitRow, cStatus), "Reviewed")
OP       RETURN(hitRow)
ERR      standard
FIXTURE  ncr_01.xlsx sheet=NcrReg cols=A:NCR No,B:Status,C:Owner rows=8 data=A:values("NCR-101"x1,"NCR-1040"x1," ncr-104 "x1,"NCR-104"x1,"NCR-105"x1,"NCR-10"x1,"NCR-106"x1,"NCR-107"x1)+boundary("NCR-101","NCR-1040"," ncr-104 ","NCR-104","NCR-105","NCR-10","NCR-106","NCR-107");B:values("Open"x6,"Closed"x2)+boundary("Open","Closed","Open","Open","Open","Closed","Open","Open");C:text(6)
FIXTURE  ncr_01.xlsx sheet=Notes cols=A:Note,B:Detail rows=2 data=A:text(8);B:text(8)
FIXTURE  ncr_02.xlsx sheet=NcrReg cols=A:NCR No,B:Status,C:Owner rows=5 data=A:values("NCR-101"x1,"NCR-1040"x1,"NCR-10"x1,"NCR-105"x1,"NCR-1004"x1)+boundary("NCR-101","NCR-1040","NCR-10","NCR-105","NCR-1004");B:values("Open"x3,"Closed"x2)+boundary("Open","Closed","Open","Open","Closed");C:text(6)
TEST     fixture=ncr_01.xlsx call=MarkNcr104Reviewed(NcrReg) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(NcrReg![Status], "Reviewed") = 1 AND VALUE(NcrReg![Status]4) = "Reviewed" AND VALUE(NcrReg![Status]5) = "Open"
TEST     fixture=ncr_01.xlsx call=MarkNcr104Reviewed(NcrReg) assert=COUNTA(NcrReg![NCR No]) = PRE_COUNTA(NcrReg![NCR No]) AND COUNTIF(NcrReg![Status], "Open") = PRE_COUNTIF(NcrReg![Status], "Open") - 1 AND COUNTIF(NcrReg![Status], "Closed") = PRE_COUNTIF(NcrReg![Status], "Closed")
TEST     fixture=ncr_02.xlsx call=MarkNcr104Reviewed(NcrReg) assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(NcrReg![Status], "Reviewed") = 0 AND COUNTA(NcrReg![Status]) = PRE_COUNTA(NcrReg![Status])
TEST     fixture=ncr_01.xlsx call=MarkNcr104Reviewed(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
