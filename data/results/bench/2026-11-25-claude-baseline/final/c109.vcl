VCL 0.7
TASK
# ASSUME: both sheets have headings in row 1; an NDE row with a blank Weld Seam No is ignored
# ASSUME: Weld Seam No values are matched with SAME (trimmed, case ignored) and Result Reject is matched with SAME
# ASSUME: the NDE Check heading is added in the next free column of the ITP Tracker when it is missing
# ASSUME: the number of points need attention is the number marked NDE missing or NDE reject
GOAL     Mark each ITP point whose weld seam has no NDE report row or has a rejected one, clear the others, and count the points marked.
PROC     CheckItpAgainstNde
KIND     Function
MODULE   modNde
ARGS     itp As Worksheet, nde As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cCheck As Long
VAR      nSeam As Long
VAR      nRes As Long
VAR      seen As Dictionary
VAR      rejected As Dictionary
VAR      seam As String
VAR      flagged As Long
PRE      COL($itp, "ITP No") > 0; COL($itp, "Weld Seam No") > 0; COL($nde, "Weld Seam No") > 0; COL($nde, "Result") > 0
MODIFIES $itp
OP       LET(nSeam, COL($nde, "Weld Seam No"))
OP       LET(nRes, COL($nde, "Result"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($nde, nSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($nde, r, nSeam))
OP         DICT_ADD(seen, TRIM(CELL($nde, r, nSeam)), 1, mode=keep)
OP         IF SAME(CELL($nde, r, nRes), "Reject")
OP           DICT_ADD(rejected, TRIM(CELL($nde, r, nSeam)), 1, mode=keep)
OP       LET(cSeam, COL($itp, "Weld Seam No"))
OP       LET(cCheck, COL($itp, "NDE Check"))
OP       IF cCheck = 0
OP         LET(cCheck, LASTCOL($itp, 1) + 1)
OP         SET_VALUE(CELL($itp, 1, cCheck), "NDE Check")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($itp, cSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($itp, r, cSeam))
OP         LET(seam, TRIM(CELL($itp, r, cSeam)))
OP         IF EXISTS(rejected, seam)
OP           SET_VALUE(CELL($itp, r, cCheck), "NDE reject")
OP           LET(flagged, flagged + 1)
OP         ELSE
OP           IF EXISTS(seen, seam)
OP             CLEAR(CELL($itp, r, cCheck), what=contents)
OP           ELSE
OP             SET_VALUE(CELL($itp, r, cCheck), "NDE missing")
OP             LET(flagged, flagged + 1)
OP       RETURN(flagged)
ERR      standard
FIXTURE  nde_01.xlsx sheet=Itp cols=A:ITP No,B:Point,C:Weld Seam No,D:Inspection Status,E:NDE Check rows=6 data=A:seq(1,1);B:text(6);C:values("W-1"x1,"W-2"x1,"W-3"x1,"W-4"x1,""x1,"w-1"x1)+boundary("W-1","W-2","W-3","W-4","","w-1");D:values("Open"x6);E:values("old"x6)
FIXTURE  nde_01.xlsx sheet=Nde cols=A:Weld Seam No,B:Method,C:Result rows=5 data=A:values("W-1"x2,"W-2"x1,"W-4"x1,""x1)+boundary("W-1","W-1","W-2","W-4","");B:text(4);C:values("Accept"x3,"Reject"x2)+boundary("Accept","Reject","Accept","Reject","Accept")
FIXTURE  nde_02.xlsx sheet=Itp cols=A:ITP No,B:Point,C:Weld Seam No,D:Inspection Status rows=6 data=A:seq(1,1);B:text(6);C:values("W-1"x1,"W-2"x1,"W-3"x1,"W-4"x1,""x1,"w-1"x1)+boundary("W-1","W-2","W-3","W-4","","w-1");D:values("Open"x6)
FIXTURE  nde_02.xlsx sheet=Nde cols=A:Weld Seam No,B:Method,C:Result rows=5 data=A:values("W-1"x2,"W-2"x1,"W-4"x1,""x1)+boundary("W-1","W-1","W-2","W-4","");B:text(4);C:values("Accept"x3,"Reject"x2)+boundary("Accept","Reject","Accept","Reject","Accept")
TEST     fixture=nde_01.xlsx call=CheckItpAgainstNde(Itp, Nde) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(Itp![NDE Check], "NDE reject") = 3 AND COUNTIF(Itp![NDE Check], "NDE missing") = 1 AND COUNTIF(Itp![NDE Check], "old") = 1 AND COUNTIF(Itp![NDE Check], "old") + COUNTIF(Itp![NDE Check], "NDE reject") + COUNTIF(Itp![NDE Check], "NDE missing") = 5
TEST     fixture=nde_01.xlsx call=CheckItpAgainstNde(Itp, Nde) assert=COUNTA(Itp![ITP No]) = PRE_COUNTA(Itp![ITP No]) AND COUNTA(Nde![Result]) = PRE_COUNTA(Nde![Result]) AND VALUE(Itp![NDE Check]3) = "" AND VALUE(Itp![NDE Check]6) = "old"
TEST     fixture=nde_02.xlsx call=CheckItpAgainstNde(Itp, Nde) assert=RESULT = 4 AND ERROR_COUNT = 0 AND VALUE(Itp![NDE Check]1) = "NDE Check" AND COUNTIF(Itp![NDE Check], "NDE reject") = 3 AND COUNTIF(Itp![NDE Check], "NDE missing") = 1
TEST     fixture=nde_01.xlsx call=CheckItpAgainstNde(Nde, Itp) assert=RESULT = -1 AND ERROR_COUNT = 1
