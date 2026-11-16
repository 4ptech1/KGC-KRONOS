VCL 0.6
TASK
# ASSUME: the NCR Register headers are in row 1 and include "NCR No", "Severity" and "Status"; data starts in row 2
# ASSUME: "Major" and "Open" are matched ignoring case and extra spaces, so " open" and "OPEN " count as Open
# ASSUME: the macro owns the fills of the data rows: each row is reset to no fill, then every Open Major row gets light red #FFC7CE across the header's columns
# ASSUME: the header row and the other sheets are not touched
GOAL     Colour every NCR that is Major and still Open light red, and return how many there are.
PROC     HighlightOpenMajorNcrs
KIND     Function
MODULE   modNcrReview
ARGS     ws As Worksheet
RET      Long
VAR      cNcr As Long
VAR      cSeverity As Long
VAR      cStatus As Long
VAR      lastCol As Long
VAR      n As Long
PRE      COL($ws, "NCR No") > 0; COL($ws, "Severity") > 0; COL($ws, "Status") > 0
MODIFIES $ws fills
OP       LET(cNcr, COL($ws, "NCR No"))
OP       LET(cSeverity, COL($ws, "Severity"))
OP       LET(cStatus, COL($ws, "Status"))
OP       LET(lastCol, LASTCOL($ws, 1))
OP       LET(n, 0)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cNcr) var=r
OP         FORMAT(ROW_RANGE($ws, r, 1, lastCol), fill=none)
OP         IF SAME(CELL($ws, r, cSeverity), "Major") AND SAME(CELL($ws, r, cStatus), "Open")
OP           FORMAT(ROW_RANGE($ws, r, 1, lastCol), fill=#FFC7CE)
OP           LET(n, n + 1)
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  ncr_01.xlsx sheet=NCR cols=A:NCR No,B:Severity,C:Status rows=8 stale=2 data=A:seq(1,1);B:values("Major"x3," major "x1,"Minor"x4)+boundary("Major"," major ","Major","Minor","Major","Minor","Minor","Minor");C:values("Open"x2," open"x1,"Closed"x3,"On Hold"x2)+boundary("Open"," open","Closed","Open","On Hold","Closed","Closed","On Hold")
FIXTURE  ncr_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=ncr_01.xlsx call=HighlightOpenMajorNcrs(NCR) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND FILL_COUNT(NCR![NCR No], #FFC7CE) = 2 AND FILL_COUNT(NCR![Status], #FFC7CE) = 2 AND FILL_COUNT(NCR![NCR No], #FF0000) = 0
TEST     fixture=ncr_01.xlsx call=HighlightOpenMajorNcrs(NCR) assert=RESULT = 2 AND FILL_COUNT(NCR![NCR No]2, #FFC7CE) = 1 AND FILL_COUNT(NCR![NCR No]3, #FFC7CE) = 1 AND FILL_COUNT(NCR![NCR No]4, #FFC7CE) = 0 AND FILL_COUNT(NCR![NCR No]5, #FFC7CE) = 0 AND COUNTA(NCR![Status]) = PRE_COUNTA(NCR![Status])
TEST     fixture=ncr_01.xlsx call=HighlightOpenMajorNcrs(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
