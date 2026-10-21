VCL 0.6
TASK
# ASSUME: Ncr headers in row 1 include "NCR No"
# ASSUME: NCR numbers that differ only in upper and lower case are duplicates; the first row of each stays
# ASSUME: DEDUPE does not trim its keys, so NCR numbers that differ in leading or trailing spaces count as different
# ASSUME: a row with a blank NCR No is never removed
GOAL     Remove the duplicate NCR No rows from the NCR Register, keeping the first row of each NCR number, and return how many rows were removed.
PROC     RemoveDuplicateNcrs
KIND     Function
MODULE   modNcrCleanup
ARGS     ncr As Worksheet
RET      Long
VAR      cNcr As Long
VAR      removed As Long
PRE      COL($ncr, "NCR No") > 0
MODIFIES $ncr
OP       LET(cNcr, COL($ncr, "NCR No"))
OP       DEDUPE(USED_RANGE($ncr), cols=cNcr, header=yes, into=removed)
OP       RETURN(removed)
ERR      standard
FIXTURE  ncr_01.xlsx sheet=Ncr cols=A:NCR No,B:Status,C:Line rows=10 data=A:text(8)+boundary("NCR-001","NCR-002","ncr-001","NCR-003","","NCR-002","","NCR-004"," NCR-003","Ncr-004");B:values("Open"x6,"Closed"x4);C:seq(1,1)
FIXTURE  ncr_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=ncr_01.xlsx call=RemoveDuplicateNcrs(Ncr) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND ROWS_USED(Ncr) = 8 AND SUM(Ncr![Line]) = 36
TEST     fixture=ncr_01.xlsx call=RemoveDuplicateNcrs(Ncr) assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTA(Ncr![NCR No]) = PRE_COUNTA(Ncr![NCR No]) - RESULT AND VALUE(Ncr![NCR No]2) = "NCR-001" AND VALUE(Ncr![Line]4) = 4 AND VALUE(Ncr![NCR No]8) = " NCR-003"
TEST     fixture=ncr_01.xlsx call=RemoveDuplicateNcrs(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
