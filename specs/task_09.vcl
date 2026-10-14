VCL 0.1
TASK
# ASSUME: NCR Register headers (row 1) include "NCR No", "Issue Date" and "Status"
# ASSUME: a row is copied when Status is "Open" (case and spaces ignored); blank Status rows are ignored
# ASSUME: the Summary sheet is rebuilt every run with the register's header row plus a Days Open column, oldest first
# ASSUME: rows with a missing or invalid Issue Date are still copied, marked "Issue date missing", and sort to the bottom
PROC     BuildOpenNcrSummary
KIND     Function
MODULE   modQuality
ARGS     ncr As Worksheet, out As Worksheet
RET      Long
VAR      cNcr As Long
VAR      cDate As Long
VAR      cStatus As Long
VAR      lastCol As Long
VAR      outRow As Long
PRE      COL($ncr, "NCR No") > 0; COL($ncr, "Issue Date") > 0; COL($ncr, "Status") > 0
OP       LET(cNcr, COL($ncr, "NCR No"))
OP       LET(cDate, COL($ncr, "Issue Date"))
OP       LET(cStatus, COL($ncr, "Status"))
OP       LET(lastCol, LASTCOL($ncr, 1))
OP       CLEAR(USED_RANGE($out), what=all)
OP       COPY_ROW($ncr, 1, $out, 1, cols=lastCol)
OP       SET_VALUE(CELL($out, 1, lastCol + 1), "Days Open")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, cNcr) var=r
OP         IF UPPER(TRIM(CELL($ncr, r, cStatus))) = "OPEN"
OP           COPY_ROW($ncr, r, $out, outRow, cols=lastCol)
OP           IF IS_DATE(CELL($ncr, r, cDate)) THEN SET_VALUE(CELL($out, outRow, lastCol + 1), TODAY - CELL($ncr, r, cDate))
OP           IF NOT IS_DATE(CELL($ncr, r, cDate)) THEN SET_VALUE(CELL($out, outRow, lastCol + 1), "Issue date missing")
OP           LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=cDate, order=asc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  ncr_01.xlsx sheet=NCR cols=A:NCR No,B:Issue Date,C:Description,D:Owner,E:Status rows=30 data=A:seq(1,1);B:dates(2026-01-01,2026-09-30)+blank(2);C:text(20);D:values("QA"x10,"Production"x10,"Stores"x10);E:values("Open"x12,"Closed"x15,""x3)
FIXTURE  ncr_01.xlsx sheet=NcrSummary cols=A:NCR No rows=0
TEST     fixture=ncr_01.xlsx call=BuildOpenNcrSummary(NCR, NcrSummary) runs=2 assert=RESULT = 12 AND COUNTA(NcrSummary![NCR No]) = 13 AND ERROR_COUNT = 0
TEST     fixture=ncr_01.xlsx call=BuildOpenNcrSummary(NCR, NcrSummary) assert=COUNTIF(NcrSummary![Status], "Closed") = 0 AND VALUE(NcrSummary![Issue Date]2) <= VALUE(NcrSummary![Issue Date]3)
TEST     fixture=ncr_01.xlsx call=BuildOpenNcrSummary(NCR, NcrSummary) assert=COUNTA(NCR![Status]) = PRE_COUNTA(NCR![Status]) AND COUNTA(NCR![Issue Date]) = PRE_COUNTA(NCR![Issue Date])
TEST     fixture=ncr_01.xlsx call=BuildOpenNcrSummary(NcrSummary, NcrSummary) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(NcrSummary![NCR No]) = PRE_COUNTA(NcrSummary![NCR No])
