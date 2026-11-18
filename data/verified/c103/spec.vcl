VCL 0.6
TASK
# ASSUME: NDE Report headers in row 1 include "Weld Seam No", "Method" and "Result"
# ASSUME: each row with a method counts as one weld tested; RT and rt are one method, shown in capitals
# ASSUME: a Result of Reject is matched ignoring case and extra spaces; rows with a blank Method are ignored
# ASSUME: the Method Summary sheet exists, is passed in and is rebuilt on every run, methods sorted A to Z
GOAL     Rebuild the Method Summary sheet with each NDE method, how many welds were tested and how many were rejected, sorted A to Z.
PROC     SummariseNdeMethods
KIND     Function
MODULE   modNde
ARGS     nde As Worksheet, out As Worksheet
RET      Long
VAR      cMethod As Long
VAR      cResult As Long
VAR      tested As Dictionary
VAR      rejected As Dictionary
VAR      nT As Long
VAR      nR As Long
VAR      outRow As Long
PRE      COL($nde, "Method") > 0; COL($nde, "Result") > 0
MODIFIES $out
OP       LET(cMethod, COL($nde, "Method"))
OP       LET(cResult, COL($nde, "Result"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Method", "Welds Tested", "Rejects")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($nde, cMethod) var=r
OP         SKIP_IF IS_EMPTY(CELL($nde, r, cMethod))
OP         DICT_ADD(tested, UPPER(TRIM(CELL($nde, r, cMethod))), 1, mode=sum)
OP         DICT_ADD(rejected, UPPER(TRIM(CELL($nde, r, cMethod))), 0, mode=sum)
OP         IF SAME(CELL($nde, r, cResult), "Reject")
OP           DICT_ADD(rejected, UPPER(TRIM(CELL($nde, r, cMethod))), 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(tested)
OP         DICT_GET(tested, k, into=nT)
OP         DICT_GET(rejected, k, into=nR)
OP         WRITE_ROW($out, outRow, k, nT, nR)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       RETURN(COUNT(tested))
PERF     screen_off
ERR      standard
FIXTURE  nde_01.xlsx sheet=NdeReport cols=A:Weld Seam No,B:Method,C:Result rows=12 data=A:seq(1,1);B:values("UT"x3,"RT"x2,"rt"x1," RT "x1,"MT"x2,"PT"x1,""x2)+boundary("UT","RT","rt","MT","","PT","UT"," RT ","","MT","UT","RT");C:values("Reject"x3,"REJECT "x1,"Pass"x6,"Repair"x1,""x1)+boundary("Reject","Pass","Reject","Pass","Reject","Pass","REJECT ","Pass","Repair","Pass","Pass","")
FIXTURE  nde_01.xlsx sheet=Summary cols=A:Method,B:Welds Tested,C:Rejects rows=0 stale=2
TEST     fixture=nde_01.xlsx call=SummariseNdeMethods(NdeReport, Summary) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(Summary![Method]) = 5 AND COUNTIF(Summary![Method], "stale") = 0 AND VALUE(Summary![Method]2) = "MT" AND VALUE(Summary![Welds Tested]2) = 2 AND VALUE(Summary![Rejects]2) = 0 AND VALUE(Summary![Method]3) = "PT" AND VALUE(Summary![Welds Tested]3) = 1 AND VALUE(Summary![Rejects]3) = 0 AND VALUE(Summary![Method]4) = "RT" AND VALUE(Summary![Welds Tested]4) = 4 AND VALUE(Summary![Rejects]4) = 1 AND VALUE(Summary![Method]5) = "UT" AND VALUE(Summary![Welds Tested]5) = 3 AND VALUE(Summary![Rejects]5) = 2
TEST     fixture=nde_01.xlsx call=SummariseNdeMethods(NdeReport, Summary) assert=SUM(Summary![Welds Tested]) = 10 AND SUM(Summary![Rejects]) = 3 AND COUNTA(NdeReport![Method]) = PRE_COUNTA(NdeReport![Method]) AND COUNTIF(NdeReport![Result], "Reject") = PRE_COUNTIF(NdeReport![Result], "Reject") AND ROWS_USED(NdeReport) = PRE_ROWS_USED(NdeReport)
TEST     fixture=nde_01.xlsx call=SummariseNdeMethods(Summary, NdeReport) assert=RESULT = -1 AND ERROR_COUNT = 1
