VCL 0.6
TASK
# ASSUME: NDE Report headers in row 1 include "Inspector" and "Result"
# ASSUME: Inspector Summary is an existing sheet that is rebuilt on every run
# ASSUME: inspectors are matched ignoring case and stray spaces, and shown as first typed, trimmed
# ASSUME: a report counts as a Reject when Result is Reject, ignoring case and spaces
GOAL     Rebuild the Inspector Summary sheet with each inspector, how many NDE reports they signed and how many were Reject, A to Z.
PROC     BuildInspectorSummary
KIND     Function
MODULE   modNde
ARGS     nde As Worksheet, out As Worksheet
RET      Long
VAR      cInsp As Long
VAR      cResult As Long
VAR      signed As Dictionary
VAR      rejects As Dictionary
VAR      nSigned As Long
VAR      nRej As Long
VAR      outRow As Long
PRE      COL($nde, "Inspector") > 0; COL($nde, "Result") > 0
MODIFIES $out
OP       LET(cInsp, COL($nde, "Inspector"))
OP       LET(cResult, COL($nde, "Result"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Inspector", "Reports Signed", "Rejects")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($nde, cInsp) var=r
OP         SKIP_IF IS_EMPTY(CELL($nde, r, cInsp))
OP         DICT_ADD(signed, TRIM(CELL($nde, r, cInsp)), 1, mode=sum)
OP         IF SAME(CELL($nde, r, cResult), "Reject")
OP           DICT_ADD(rejects, TRIM(CELL($nde, r, cInsp)), 1, mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(signed)
OP         DICT_GET(signed, k, into=nSigned)
OP         LET(nRej, 0)
OP         IF EXISTS(rejects, k)
OP           DICT_GET(rejects, k, into=nRej)
OP         WRITE_ROW($out, outRow, k, nSigned, nRej)
OP         LET(outRow, outRow + 1)
OP       IF outRow > 2
OP         SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  nde_01.xlsx sheet=Nde cols=A:Weld Seam No,B:Inspector,C:Result rows=7 data=A:seq(1,1);B:values("Zed Park"x2,"Amy Cole"x2,"amy cole "x1,""x1,"Ben Ito"x1)+boundary("Zed Park","Amy Cole","amy cole ","","Zed Park","Ben Ito","Amy Cole");C:values("Reject"x2,"reject "x1,"Accept"x4)+boundary("Accept","Reject","Accept","Reject","Accept","reject ","Accept")
FIXTURE  nde_01.xlsx sheet=Summary cols=A:Inspector,B:Reports Signed,C:Rejects rows=0 stale=3
FIXTURE  nde_02.xlsx sheet=Nde cols=A:Weld Seam No,B:Inspector,C:Result rows=0
FIXTURE  nde_02.xlsx sheet=Summary cols=A:Inspector,B:Reports Signed,C:Rejects rows=0 stale=2
TEST     fixture=nde_01.xlsx call=BuildInspectorSummary(Nde, Summary) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTA(Summary![Inspector]) = 4 AND COUNTIF(Summary![Inspector], "stale") = 0 AND VALUE(Summary![Inspector]2) = "Amy Cole" AND VALUE(Summary![Reports Signed]2) = 3 AND VALUE(Summary![Rejects]2) = 1 AND VALUE(Summary![Inspector]3) = "Ben Ito" AND VALUE(Summary![Reports Signed]3) = 1 AND VALUE(Summary![Rejects]3) = 1 AND VALUE(Summary![Inspector]4) = "Zed Park" AND VALUE(Summary![Reports Signed]4) = 2 AND VALUE(Summary![Rejects]4) = 0
TEST     fixture=nde_01.xlsx call=BuildInspectorSummary(Nde, Summary) assert=ERROR_COUNT = 0 AND SUM(Summary![Reports Signed]) = 6 AND SUM(Summary![Rejects]) = 2 AND COUNTA(Nde![Weld Seam No]) = PRE_COUNTA(Nde![Weld Seam No]) AND COUNTA(Nde![Inspector]) = PRE_COUNTA(Nde![Inspector]) AND COUNTIF(Nde![Result], "Reject") = PRE_COUNTIF(Nde![Result], "Reject")
TEST     fixture=nde_02.xlsx call=BuildInspectorSummary(Nde, Summary) assert=RESULT = 0 AND ERROR_COUNT = 0 AND ROWS_USED(Summary) = 1 AND COUNTA(Nde![Weld Seam No]) = PRE_COUNTA(Nde![Weld Seam No])
TEST     fixture=nde_01.xlsx call=BuildInspectorSummary(Summary, Nde) assert=RESULT = -1 AND ERROR_COUNT = 1
