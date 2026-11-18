VCL 0.6
TASK
# ASSUME: NCR Register headers in row 1 include "NCR No", "Status" and "Disposition"
# ASSUME: Closed is matched ignoring case and extra spaces; a Disposition of only spaces counts as none
# ASSUME: the old fill is removed from every NCR No cell first, so a re-run gives the same result; other columns are untouched
# ASSUME: the last row is found from NCR No
GOAL     Remove the fill from every NCR No cell, then colour orange the NCR No of each Closed NCR with no Disposition and return how many were coloured.
PROC     ColourClosedWithoutDisposition
KIND     Function
MODULE   modNcr
ARGS     ws As Worksheet
RET      Long
VAR      cNcr As Long
VAR      cStatus As Long
VAR      cDisp As Long
VAR      coloured As Long
PRE      COL($ws, "NCR No") > 0; COL($ws, "Status") > 0; COL($ws, "Disposition") > 0
MODIFIES $ws fills
OP       LET(cNcr, COL($ws, "NCR No"))
OP       LET(cStatus, COL($ws, "Status"))
OP       LET(cDisp, COL($ws, "Disposition"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cNcr) var=r
OP         FORMAT(CELL($ws, r, cNcr), fill=none)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cNcr) var=r
OP         IF SAME(CELL($ws, r, cStatus), "Closed")
OP           IF IS_EMPTY(CELL($ws, r, cDisp))
OP             FORMAT(CELL($ws, r, cNcr), fill=#FFA500)
OP             LET(coloured, coloured + 1)
OP       RETURN(coloured)
ERR      standard
FIXTURE  ncr_01.xlsx sheet=Ncr cols=A:NCR No,B:Severity,C:Status,D:Disposition rows=10 stale=10 data=A:seq(1,1);B:values("Major"x5,"Minor"x5);C:values("Closed"x5,"Open"x4," closed "x1)+boundary("Closed","Closed","Open"," closed ","Closed","Open","Closed","Open","Open","Closed");D:values("Rework"x3,""x6," "x1)+boundary("","Rework","","","Rework","","Rework"," ","","")
FIXTURE  ncr_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=ncr_01.xlsx call=ColourClosedWithoutDisposition(Ncr) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND FILL_COUNT(Ncr![NCR No], #FFA500) = 3 AND FILL_COUNT(Ncr![NCR No], #FF0000) = 0
TEST     fixture=ncr_01.xlsx call=ColourClosedWithoutDisposition(Ncr) assert=FILL_COUNT(Ncr![NCR No]2, #FFA500) = 1 AND FILL_COUNT(Ncr![NCR No]3, #FFA500) = 0 AND FILL_COUNT(Ncr![NCR No]5, #FFA500) = 1 AND FILL_COUNT(Ncr![NCR No]6, #FFA500) = 0 AND FILL_COUNT(Ncr![NCR No]8, #FFA500) = 0 AND FILL_COUNT(Ncr![NCR No]9, #FFA500) = 0 AND FILL_COUNT(Ncr![NCR No]11, #FFA500) = 1
TEST     fixture=ncr_01.xlsx call=ColourClosedWithoutDisposition(Ncr) assert=COUNTA(Ncr![NCR No]) = PRE_COUNTA(Ncr![NCR No]) AND COUNTA(Ncr![Status]) = PRE_COUNTA(Ncr![Status]) AND COUNTA(Ncr![Disposition]) = PRE_COUNTA(Ncr![Disposition])
TEST     fixture=ncr_01.xlsx call=ColourClosedWithoutDisposition(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
