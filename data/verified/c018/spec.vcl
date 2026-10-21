VCL 0.6
TASK
# ASSUME: NDE Report headers in row 1 include "Method" and "Result"
# ASSUME: the real sheets are "NDE Report" and "Method Summary"; the TESTs use Nde and MethodSummary
# ASSUME: Method Summary already exists and is rebuilt on every run, with the headers Method, Accepted, Repair and No Result
# ASSUME: each row is one seam tested by one method, so rows are counted; Method is trimmed and matched ignoring case
# ASSUME: a blank Result means no result yet; Accepted and Repair are matched ignoring case and extra spaces
# ASSUME: any other Result (such as Rejected) is not counted in any column; a row with a blank Method is skipped
# ASSUME: methods are listed in the order they first appear in the report
GOAL     Count, for each NDE method, the seams that are Accepted, the seams that need Repair and the seams with no result yet, and write them on the Method Summary sheet.
PROC     SummariseNdeByMethod
KIND     Function
MODULE   modNde
ARGS     nde As Worksheet, out As Worksheet
RET      Long
VAR      cMethod As Long
VAR      cResult As Long
VAR      accepted As Dictionary
VAR      repair As Dictionary
VAR      pending As Dictionary
VAR      aCount As Long
VAR      rCount As Long
VAR      pCount As Long
VAR      outRow As Long
PRE      COL($nde, "Method") > 0; COL($nde, "Result") > 0
MODIFIES $out
OP       LET(cMethod, COL($nde, "Method"))
OP       LET(cResult, COL($nde, "Result"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Method", "Accepted", "Repair", "No Result")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($nde, cMethod) var=r
OP         SKIP_IF IS_EMPTY(CELL($nde, r, cMethod))
OP         DICT_ADD(accepted, TRIM(CELL($nde, r, cMethod)), IIF(SAME(CELL($nde, r, cResult), "Accepted"), 1, 0), mode=sum)
OP         DICT_ADD(repair, TRIM(CELL($nde, r, cMethod)), IIF(SAME(CELL($nde, r, cResult), "Repair"), 1, 0), mode=sum)
OP         DICT_ADD(pending, TRIM(CELL($nde, r, cMethod)), IIF(IS_EMPTY(CELL($nde, r, cResult)), 1, 0), mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(accepted)
OP         DICT_GET(accepted, k, into=aCount)
OP         DICT_GET(repair, k, into=rCount)
OP         DICT_GET(pending, k, into=pCount)
OP         WRITE_ROW($out, outRow, k, aCount, rCount, pCount)
OP         LET(outRow, outRow + 1)
OP       RETURN(COUNT(accepted))
PERF     screen_off
ERR      standard
FIXTURE  nde_report_01.xlsx sheet=Nde cols=A:Weld Seam No,B:Method,C:Result rows=16 data=A:seq(1,1);B:values("RT"x5," rt "x1,"UT"x4,"MT"x3,"PT"x2,""x1)+boundary("RT","UT","RT","MT"," rt ","UT","","PT","RT","UT","MT","RT","UT","PT","MT","RT");C:values("Accepted"x6," accepted "x1,"Repair"x4,""x4,"Rejected"x1)+boundary("Accepted","Repair","","Accepted","Repair","Accepted","Accepted","","Accepted",""," accepted ","Rejected","Repair","Repair","","Accepted")
FIXTURE  nde_report_01.xlsx sheet=MethodSummary cols=A:Method,B:Accepted,C:Repair,D:No Result rows=0 stale=2
TEST     fixture=nde_report_01.xlsx call=SummariseNdeByMethod(Nde, MethodSummary) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND ROWS_USED(MethodSummary) = 5 AND COUNTIF(MethodSummary![Method], "stale") = 0 AND VALUE(MethodSummary![Method]2) = "RT" AND VALUE(MethodSummary![Accepted]2) = 3 AND VALUE(MethodSummary![Repair]2) = 1 AND VALUE(MethodSummary![No Result]2) = 1 AND VALUE(MethodSummary![Method]3) = "UT" AND VALUE(MethodSummary![Accepted]3) = 1 AND VALUE(MethodSummary![Repair]3) = 2 AND VALUE(MethodSummary![No Result]3) = 1
TEST     fixture=nde_report_01.xlsx call=SummariseNdeByMethod(Nde, MethodSummary) assert=RESULT = 4 AND ERROR_COUNT = 0 AND VALUE(MethodSummary![Method]4) = "MT" AND VALUE(MethodSummary![Accepted]4) = 2 AND VALUE(MethodSummary![Repair]4) = 0 AND VALUE(MethodSummary![No Result]4) = 1 AND VALUE(MethodSummary![Method]5) = "PT" AND VALUE(MethodSummary![Accepted]5) = 0 AND VALUE(MethodSummary![Repair]5) = 1 AND VALUE(MethodSummary![No Result]5) = 1
TEST     fixture=nde_report_01.xlsx call=SummariseNdeByMethod(Nde, MethodSummary) assert=SUM(MethodSummary![Accepted]) = 6 AND SUM(MethodSummary![Repair]) = 4 AND SUM(MethodSummary![No Result]) = 4 AND COUNTA(Nde![Method]) = PRE_COUNTA(Nde![Method]) AND COUNTA(Nde![Result]) = PRE_COUNTA(Nde![Result]) AND COUNTIF(Nde![Result], "Repair") = PRE_COUNTIF(Nde![Result], "Repair")
TEST     fixture=nde_report_01.xlsx call=SummariseNdeByMethod(MethodSummary, Nde) assert=RESULT = -1 AND ERROR_COUNT = 1
