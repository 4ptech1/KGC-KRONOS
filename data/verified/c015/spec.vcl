VCL 0.6
TASK
# ASSUME: NCR headers in row 1 include "NCR No", "Issue Date" and "Status"
# ASSUME: still Open means Status is Open, ignoring case and extra spaces
# ASSUME: more than 30 days ago means an Issue Date before TODAY - 30; a missing or non-date Issue Date is not counted
# ASSUME: the NCR No cells carry no fills but these highlights, so each run clears every NCR No fill first
# ASSUME: a row with a blank NCR No is skipped
GOAL     Highlight in yellow the NCR No of every Open NCR issued more than 30 days ago, clearing old highlights first, and return how many were highlighted.
PROC     HighlightOverdueNcrs
KIND     Function
MODULE   modNcr
ARGS     ws As Worksheet
RET      Long
VAR      cNo As Long
VAR      cIssue As Long
VAR      cStatus As Long
VAR      marked As Long
PRE      COL($ws, "NCR No") > 0; COL($ws, "Issue Date") > 0; COL($ws, "Status") > 0
MODIFIES $ws fills
OP       LET(cNo, COL($ws, "NCR No"))
OP       LET(cIssue, COL($ws, "Issue Date"))
OP       LET(cStatus, COL($ws, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cNo) var=r
OP         FORMAT(CELL($ws, r, cNo), fill=none)
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cNo))
OP         SKIP_IF NOT SAME(CELL($ws, r, cStatus), "Open")
OP         IF IS_DATE(CELL($ws, r, cIssue))
OP           IF CELL($ws, r, cIssue) < TODAY - 30
OP             FORMAT(CELL($ws, r, cNo), fill=yellow)
OP             LET(marked, marked + 1)
OP       RETURN(marked)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  ncr_01.xlsx sheet=Ncr cols=A:NCR No,B:Raised By,C:Issue Date,D:Category,E:Status rows=9 stale=3 data=A:text(5)+boundary("N-01","N-02","N-03","N-04","N-05","","N-06","N-07","N-08");B:text(6);C:reldates(-120,-10)+boundary(-31,-30,-29,-90,"",-60,-100,-45,-31);D:values("Weld"x5,"Material"x4);E:values("Open"x6," OPEN "x1,"Closed"x2)+boundary("Open","Open","Open","Closed","Open","Open"," OPEN ","Open","Closed")
FIXTURE  ncr_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=ncr_01.xlsx call=HighlightOverdueNcrs(Ncr) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND FILL_COUNT(Ncr![NCR No], #FFFF00) = 3 AND FILL_COUNT(Ncr![NCR No], #FF0000) = 0
TEST     fixture=ncr_01.xlsx call=HighlightOverdueNcrs(Ncr) assert=RESULT = 3 AND COUNTA(Ncr![Status]) = PRE_COUNTA(Ncr![Status]) AND COUNTA(Ncr![Issue Date]) = PRE_COUNTA(Ncr![Issue Date]) AND FILL_COUNT(Ncr![NCR No]2, #FFFF00) = 1 AND FILL_COUNT(Ncr![NCR No]3, #FFFF00) = 0 AND FILL_COUNT(Ncr![NCR No]4, #FFFF00) = 0 AND FILL_COUNT(Ncr![NCR No]5, #FFFF00) = 0 AND FILL_COUNT(Ncr![NCR No]6, #FFFF00) = 0 AND FILL_COUNT(Ncr![NCR No]7, #FFFF00) = 0 AND FILL_COUNT(Ncr![NCR No]8, #FFFF00) = 1 AND FILL_COUNT(Ncr![NCR No]9, #FFFF00) = 1 AND FILL_COUNT(Ncr![NCR No]10, #FFFF00) = 0
TEST     fixture=ncr_01.xlsx call=HighlightOverdueNcrs(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
