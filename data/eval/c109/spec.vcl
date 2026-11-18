VCL 0.6
TASK
# ASSUME: ITP Tracker headers in row 1 include "Weld Seam No"; NDE Report headers include "Weld Seam No" and "Result"
# ASSUME: the real sheets are "ITP Tracker" and "NDE Report"; the TESTs use Itp and NdeReport
# ASSUME: if the NDE Check heading is missing it is added in the first empty column right of the last heading in row 1
# ASSUME: Weld Seam No values are trimmed and compared ignoring case; Reject is matched ignoring case and spaces
# ASSUME: an NDE Report row with a blank Weld Seam No is ignored; an ITP point with a blank Weld Seam No keeps whatever its NDE Check holds
# ASSUME: a point needs attention when it is marked NDE missing or NDE reject; NDE missing wins over NDE reject because a missing seam has no rows
GOAL     Mark each ITP point whose weld seam has no NDE row as NDE missing or has an NDE reject as NDE reject, clear it on the others and return how many points need attention.
PROC     CheckItpAgainstNde
KIND     Function
MODULE   modItp
ARGS     itp As Worksheet, nde As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cNdeSeam As Long
VAR      cNdeResult As Long
VAR      cCheck As Long
VAR      rejects As Dictionary
VAR      seam As String
VAR      rej As Long
VAR      attention As Long
PRE      COL($itp, "Weld Seam No") > 0; COL($nde, "Weld Seam No") > 0; COL($nde, "Result") > 0
MODIFIES $itp
OP       LET(cSeam, COL($itp, "Weld Seam No"))
OP       LET(cNdeSeam, COL($nde, "Weld Seam No"))
OP       LET(cNdeResult, COL($nde, "Result"))
OP       LET(cCheck, COL($itp, "NDE Check"))
OP       IF cCheck = 0
OP         LET(cCheck, LASTCOL($itp, 1) + 1)
OP         SET_VALUE(CELL($itp, 1, cCheck), "NDE Check")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($nde, cNdeSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($nde, r, cNdeSeam))
OP         DICT_ADD(rejects, TRIM(CELL($nde, r, cNdeSeam)), IIF(SAME(CELL($nde, r, cNdeResult), "Reject"), 1, 0), mode=sum)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($itp, cSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($itp, r, cSeam))
OP         LET(seam, TRIM(CELL($itp, r, cSeam)))
OP         IF NOT EXISTS(rejects, seam)
OP           SET_VALUE(CELL($itp, r, cCheck), "NDE missing")
OP           LET(attention, attention + 1)
OP         ELSE
OP           DICT_GET(rejects, seam, into=rej)
OP           IF rej > 0
OP             SET_VALUE(CELL($itp, r, cCheck), "NDE reject")
OP             LET(attention, attention + 1)
OP           ELSE
OP             CLEAR(CELL($itp, r, cCheck), what=contents)
OP       MSG("Points needing attention: " & attention)
OP       RETURN(attention)
PERF     screen_off
ERR      standard
FIXTURE  itp_nde_01.xlsx sheet=Itp cols=A:ITP No,B:Point,C:Weld Seam No,D:Inspection Status rows=8 data=A:seq(1,1);B:text(6);C:values("S-1"x1,"S-2"x1,"S-3"x1,""x1," S-1 "x1,"s-2"x1,"S-4"x1,"S-5"x1)+boundary("S-1","S-2","S-3",""," S-1 ","s-2","S-4","S-5");D:values("Open"x4,"Closed"x4)
FIXTURE  itp_nde_01.xlsx sheet=NdeReport cols=A:Weld Seam No,B:Method,C:Result rows=7 data=A:values("S-1"x2,"S-2"x2,"S-4"x1,""x1,"S-6"x1)+boundary("S-4","S-2","S-2","S-1","S-1","","S-6");B:values("RT"x4,"UT"x3);C:values("Pass"x3,"Reject"x3,"reject "x1)+boundary("reject ","Reject","Pass","Pass","Pass","Reject","Reject")
FIXTURE  itp_nde_02.xlsx sheet=Itp cols=A:ITP No,B:Point,C:Weld Seam No,D:Inspection Status,E:NDE Check rows=4 data=A:seq(1,1);B:text(6);C:values("S-1"x1,"S-2"x1,""x1,"S-9"x1)+boundary("S-1","S-2","","S-9");D:values("Open"x2,"Closed"x2);E:values("NDE missing"x1,""x1,"keep"x1,"NDE reject"x1)+boundary("NDE missing","","keep","NDE reject")
FIXTURE  itp_nde_02.xlsx sheet=NdeReport cols=A:Weld Seam No,B:Method,C:Result rows=7 data=A:values("S-1"x2,"S-2"x2,"S-4"x1,""x1,"S-6"x1)+boundary("S-4","S-2","S-2","S-1","S-1","","S-6");B:values("RT"x4,"UT"x3);C:values("Pass"x3,"Reject"x3,"reject "x1)+boundary("reject ","Reject","Pass","Pass","Pass","Reject","Reject")
TEST     fixture=itp_nde_01.xlsx call=CheckItpAgainstNde(Itp, NdeReport) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTIF(Itp![NDE Check], "NDE missing") = 2 AND COUNTIF(Itp![NDE Check], "NDE reject") = 3 AND COUNTA(Itp![NDE Check]) = 6 AND COUNTA(Itp![Point]) = PRE_COUNTA(Itp![Point])
TEST     fixture=itp_nde_01.xlsx call=CheckItpAgainstNde(Itp, NdeReport) assert=VALUE(Itp!E1) = "NDE Check" AND VALUE(Itp![NDE Check]2) = "" AND VALUE(Itp![NDE Check]3) = "NDE reject" AND VALUE(Itp![NDE Check]4) = "NDE missing" AND VALUE(Itp![NDE Check]5) = "" AND VALUE(Itp![NDE Check]6) = "" AND VALUE(Itp![NDE Check]7) = "NDE reject" AND VALUE(Itp![NDE Check]8) = "NDE reject" AND VALUE(Itp![NDE Check]9) = "NDE missing"
TEST     fixture=itp_nde_01.xlsx call=CheckItpAgainstNde(Itp, NdeReport) assert=ERROR_COUNT = 0 AND COUNTA(NdeReport![Weld Seam No]) = PRE_COUNTA(NdeReport![Weld Seam No]) AND COUNTA(Itp![Weld Seam No]) = PRE_COUNTA(Itp![Weld Seam No]) AND COUNTA(Itp![Inspection Status]) = PRE_COUNTA(Itp![Inspection Status])
TEST     fixture=itp_nde_02.xlsx call=CheckItpAgainstNde(Itp, NdeReport) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(Itp![NDE Check]1) = "NDE Check" AND VALUE(Itp![NDE Check]2) = "" AND VALUE(Itp![NDE Check]3) = "NDE reject" AND VALUE(Itp![NDE Check]4) = "keep" AND VALUE(Itp![NDE Check]5) = "NDE missing" AND COUNTA(Itp![NDE Check]) = 4 AND COUNTA(Itp![Point]) = PRE_COUNTA(Itp![Point])
TEST     fixture=itp_nde_01.xlsx call=CheckItpAgainstNde(NdeReport, Itp) assert=RESULT = -1 AND ERROR_COUNT = 1
