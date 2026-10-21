VCL 0.5
BUILD
# ASSUME: NCR headers (row 1) include "NCR No", "Status" and "Days Open"
# ASSUME: an Open NCR older than the limit becomes Overdue; Summary is rebuilt on every run
NAME     NcrReview
TARGET   xlsm
ENTRY    modMain.Run
MODULE   modMain
MODULE   modSteps
SIG      modMain.Run()
SIG      modMain.RunHeadless(ncr As Worksheet, out As Worksheet) As Long
SIG      modSteps.MarkOverdue(ncr As Worksheet, limit As Long) As Long
SIG      modSteps.WriteSummary(ncr As Worksheet, out As Worksheet)
FLOW     modMain.Run -> modMain.RunHeadless -> modSteps.MarkOverdue
FLOW     modMain.RunHeadless -> modSteps.WriteSummary
FIXTURE  ncr_01.xlsx sheet=NCR cols=A:NCR No,B:Status,C:Days Open rows=8 data=A:seq(1,1);B:values("Open"x5,"Closed"x2,""x1)+boundary("Open","Open","Closed","","Open");C:int(1,60)+boundary(31,30,45,10)
FIXTURE  ncr_01.xlsx sheet=Summary cols=A:Status,B:Count rows=0 stale=2
TEST     fixture=ncr_01.xlsx call=modMain.RunHeadless(NCR, Summary) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(NCR![Status], "Overdue") = 4 AND COUNTIF(NCR![Status], "Open") = 1 AND COUNTA(Summary![Status]) = 4 AND VALUE(Summary![Count]2) = 4 AND COUNTIF(Summary![Status], "stale") = 0
TEST     fixture=ncr_01.xlsx call=modMain.RunHeadless(NCR, Summary) assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(NCR![NCR No]) = PRE_COUNTA(NCR![NCR No]) AND VALUE(NCR![Status]2) = "Overdue" AND VALUE(NCR![Status]3) = "Open" AND VALUE(Summary![Status]2) = "Overdue"
TEST     fixture=ncr_01.xlsx call=modMain.RunHeadless(Summary, NCR) assert=RESULT = -1 AND ERROR_COUNT = 1
---
VCL 0.5
TASK
PROC     Run
KIND     Sub
MODULE   modMain
VAR      n As Long
OP       CALL(modMain.RunHeadless, SHEET(THIS_BOOK, "NCR Register"), SHEET(THIS_BOOK, "NCR Summary"), into=n)
OP       MSG("NCRs marked Overdue: " & n)
ERR      standard
---
VCL 0.5
TASK
PROC     RunHeadless
KIND     Function
MODULE   modMain
ARGS     ncr As Worksheet, out As Worksheet
RET      Long
VAR      marked As Long
PRE      COL($ncr, "Status") > 0; COL($ncr, "Days Open") > 0
MODIFIES $ncr["Status"]
MODIFIES $out
OP       CALL(modSteps.MarkOverdue, $ncr, 30, into=marked)
OP       CALL(modSteps.WriteSummary, $ncr, $out)
OP       RETURN(marked)
ERR      standard
---
VCL 0.5
TASK
PROC     MarkOverdue
KIND     Function
MODULE   modSteps
ARGS     ncr As Worksheet, limit As Long
RET      Long
VAR      cStatus As Long
VAR      cDays As Long
VAR      n As Long
PRE      COL($ncr, "Status") > 0; COL($ncr, "Days Open") > 0
MODIFIES $ncr["Status"]
OP       LET(cStatus, COL($ncr, "Status"))
OP       LET(cDays, COL($ncr, "Days Open"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, cStatus) var=r
OP         IF SAME(CELL($ncr, r, cStatus), "Open")
OP           IF IS_NUMBER(CELL($ncr, r, cDays))
OP             IF CELL($ncr, r, cDays) > $limit
OP               SET_VALUE(CELL($ncr, r, cStatus), "Overdue")
OP               LET(n, n + 1)
OP       RETURN(n)
ERR      standard
---
VCL 0.5
TASK
PROC     WriteSummary
KIND     Sub
MODULE   modSteps
ARGS     ncr As Worksheet, out As Worksheet
VAR      cStatus As Long
VAR      counts As Dictionary
VAR      got As Long
VAR      outRow As Long
PRE      COL($ncr, "Status") > 0
MODIFIES $out
OP       LET(cStatus, COL($ncr, "Status"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Status", "Count")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, cStatus) var=r
OP         SKIP_IF IS_EMPTY(CELL($ncr, r, cStatus))
OP         DICT_ADD(counts, TRIM(CELL($ncr, r, cStatus)), 1, mode=count)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         WRITE_ROW($out, outRow, k, got)
OP         LET(outRow, outRow + 1)
ERR      standard
