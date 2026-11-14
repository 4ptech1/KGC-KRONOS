VCL 0.6
TASK
# ASSUME: the log's headers in row 1 include "Job ID" and "Qty Issued"; the real sheets are "Consumable Issue" and "Job Issue Summary", and the TESTs use IssueLog and Summary
# ASSUME: the Job Issue Summary sheet already exists and holds nothing else: it is cleared and rebuilt on every run with the headings Job ID, Issues, Total Qty Issued
# ASSUME: Job IDs are matched ignoring case and spaces at the ends, and shown as first typed
# ASSUME: a row with a Job ID whose Qty Issued is blank or not a number still counts as one issue and adds nothing to the total
# ASSUME: jobs with the same total may come in either order; the Function returns the number of jobs listed
GOAL     Rebuild the Job Issue Summary sheet from the Consumable Issue log with one row per Job ID showing its number of issues and total Qty Issued, the biggest total first.
PROC     BuildJobIssueSummary
KIND     Function
MODULE   modConsumables
ARGS     log As Worksheet, out As Worksheet
RET      Long
VAR      cJob As Long
VAR      cQty As Long
VAR      job As String
VAR      counts As Dictionary
VAR      totals As Dictionary
VAR      issues As Long
VAR      total As Double
VAR      outRow As Long
PRE      COL($log, "Job ID") > 0; COL($log, "Qty Issued") > 0
MODIFIES $out
OP       LET(cJob, COL($log, "Job ID"))
OP       LET(cQty, COL($log, "Qty Issued"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Job ID", "Issues", "Total Qty Issued")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($log, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($log, r, cJob))
OP         LET(job, TRIM(CELL($log, r, cJob)))
OP         DICT_ADD(counts, job, 1, mode=sum)
OP         IF IS_NUMBER(CELL($log, r, cQty)) THEN DICT_ADD(totals, job, CELL($log, r, cQty), mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=issues)
OP         LET(total, 0)
OP         IF EXISTS(totals, k) THEN DICT_GET(totals, k, into=total)
OP         WRITE_ROW($out, outRow, k, issues, total)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=3, order=desc, header=yes)
OP       RETURN(COUNT(counts))
PERF     screen_off
ERR      standard
FIXTURE  issues_01.xlsx sheet=IssueLog cols=A:Issue No,B:Batch No,C:Qty Issued,D:Job ID rows=12 data=A:seq(5001,1);B:text(6);C:int(1,20)+boundary(10,5,7,100,"n/a",20,"",4,50,1,"x",40);D:values("J-1"x2," J-1 "x1,"J-2"x3,"j-2"x1,"J-3"x2,"J-4"x1,""x2)+boundary("J-1","J-2"," J-1 ","","J-3","j-2","J-1","J-2","","J-2","J-4","J-3")
FIXTURE  issues_01.xlsx sheet=Summary cols=A:Job ID,B:Issues,C:Total Qty Issued rows=0 stale=3
FIXTURE  issues_01.xlsx sheet=Returns cols=A:Return No,B:Qty Issued rows=2 data=A:seq(1,1);B:int(1,5)
FIXTURE  issues_02.xlsx sheet=IssueLog cols=A:Issue No,B:Batch No,C:Qty Issued,D:Job ID rows=0
FIXTURE  issues_02.xlsx sheet=Summary cols=A:Job ID,B:Issues,C:Total Qty Issued rows=0 stale=2
TEST     fixture=issues_01.xlsx call=BuildJobIssueSummary(IssueLog, Summary) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND ROWS_USED(Summary) = 5 AND COUNTIF(Summary![Job ID], "stale") = 0 AND FILL_COUNT(Summary![Job ID], #FF0000) = 0 AND VALUE(Summary![Job ID]2) = "J-3" AND VALUE(Summary![Issues]2) = 2 AND VALUE(Summary![Total Qty Issued]2) = 40 AND VALUE(Summary![Job ID]3) = "J-2" AND VALUE(Summary![Issues]3) = 4 AND VALUE(Summary![Total Qty Issued]3) = 30 AND VALUE(Summary![Job ID]4) = "J-1" AND VALUE(Summary![Issues]4) = 3 AND VALUE(Summary![Total Qty Issued]4) = 17 AND VALUE(Summary![Job ID]5) = "J-4" AND VALUE(Summary![Issues]5) = 1 AND VALUE(Summary![Total Qty Issued]5) = 0
TEST     fixture=issues_01.xlsx call=BuildJobIssueSummary(IssueLog, Summary) assert=RESULT = 4 AND COUNTA(IssueLog![Job ID]) = PRE_COUNTA(IssueLog![Job ID]) AND COUNTA(IssueLog![Qty Issued]) = PRE_COUNTA(IssueLog![Qty Issued]) AND SUM(IssueLog![Qty Issued]) = PRE_SUM(IssueLog![Qty Issued]) AND SUM(Summary![Issues]) = 10 AND SUM(Summary![Total Qty Issued]) = 87
TEST     fixture=issues_01.xlsx call=BuildJobIssueSummary(Summary, IssueLog) assert=RESULT = -1 AND ERROR_COUNT = 1
TEST     fixture=issues_01.xlsx call=BuildJobIssueSummary(Returns, Summary) assert=RESULT = -1 AND ERROR_COUNT = 1
TEST     fixture=issues_02.xlsx call=BuildJobIssueSummary(IssueLog, Summary) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND ROWS_USED(Summary) = 1 AND COUNTA(Summary![Job ID]) = 1
