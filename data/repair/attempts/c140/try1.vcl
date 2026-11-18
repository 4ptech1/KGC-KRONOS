VCL 0.6
TASK
# ASSUME: ITP Tracker headers in row 1 include "ITP No" and "NCR Hold"; NCR Register headers include "ITP No" and "Status"
# ASSUME: a point is on hold when any NCR Register row with the same ITP No (ignoring case and stray spaces) has a Status that is not Closed
# ASSUME: Closed is matched ignoring case and spaces; an NCR with a blank Status counts as not Closed
# ASSUME: an NCR row with no ITP No is ignored; an ITP Tracker row with no ITP No is skipped and left as it is
GOAL     Write Hold in NCR Hold for every ITP Tracker point whose ITP No has an NCR that is not Closed, clear the others, and return how many points are on hold.
PROC     MarkNcrHolds
KIND     Function
MODULE   modItpHold
ARGS     itp As Worksheet, ncr As Worksheet
RET      Long
VAR      cItp As Long
VAR      cHold As Long
VAR      nItp As Long
VAR      nStatus As Long
VAR      open As Dictionary
VAR      n As Long
PRE      COL($itp, "ITP No") > 0; COL($itp, "NCR Hold") > 0; COL($ncr, "ITP No") > 0; COL($ncr, "Status") > 0
MODIFIES $itp["NCR Hold"]
OP       LET(cItp, COL($itp, "ITP No"))
OP       LET(cHold, COL($itp, "NCR Hold"))
OP       LET(nItp, COL($ncr, "ITP No"))
OP       LET(nStatus, COL($ncr, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ncr, nItp) var=q
OP         SKIP_IF IS_EMPTY(CELL($ncr, q, nItp))
OP         IF NOT SAME(CELL($ncr, q, nStatus), "Closed")
OP           DICT_ADD(open, TRIM(CELL($ncr, q, nItp)), 1, mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($itp, cItp) var=r
OP         SKIP_IF IS_EMPTY(CELL($itp, r, cItp))
OP         IF EXISTS(open, TRIM(CELL($itp, r, cItp)))
OP           SET_VALUE(CELL($itp, r, cHold), "Hold")
OP           LET(n, n + 1)
OP         ELSE
OP           SET_VALUE(CELL($itp, r, cHold), "")
OP       MSG("Points on hold: " & n)
OP       RETURN(n)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  itphold_01.xlsx sheet=Itp cols=A:ITP No,B:Point,C:Inspection Status,D:NCR Hold rows=9 data=A:values("ITP-1"x2," itp-2"x1,"ITP-3"x2,"ITP-4"x1,"ITP-5"x1,""x1,"ITP-1 "x1)+boundary("ITP-1"," itp-2","ITP-3","ITP-4","","ITP-5","ITP-3","ITP-1","ITP-1 ");B:seq(1,1);C:values("Open"x9);D:values("Hold"x2,"old"x1,""x6)+boundary("","Hold","","old","keep","","Hold","","")
FIXTURE  itphold_01.xlsx sheet=Ncr cols=A:NCR No,B:ITP No,C:Status rows=7 data=A:seq(1,1);B:values("ITP-1"x1,"itp-2 "x1,"ITP-3"x2,"ITP-5"x1,""x1,"ITP-9"x1)+boundary("ITP-1","itp-2 ","ITP-3","ITP-3","","ITP-5","ITP-9");C:values("Closed"x3,"open"x1,"Open"x1," closed "x1,""x1)+boundary("Open"," CLOSED ","Closed","Closed","Open","Closed","")
TEST     fixture=itphold_01.xlsx call=MarkNcrHolds(Itp, Ncr) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Itp![NCR Hold], "Hold") = 3 AND COUNTIF(Itp![NCR Hold], "old") = 0
TEST     fixture=itphold_01.xlsx call=MarkNcrHolds(Itp, Ncr) assert=ERROR_COUNT = 0 AND COUNTA(Itp![ITP No]) = PRE_COUNTA(Itp![ITP No]) AND COUNTA(Itp![Point]) = PRE_COUNTA(Itp![Point]) AND SUM(Itp![Point]) = PRE_SUM(Itp![Point]) AND COUNTA(Ncr![Status]) = PRE_COUNTA(Ncr![Status])
TEST     fixture=itphold_01.xlsx call=MarkNcrHolds(Ncr, Itp) assert=RESULT = -1 AND ERROR_COUNT = 1
