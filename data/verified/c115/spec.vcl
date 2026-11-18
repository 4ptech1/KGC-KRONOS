VCL 0.6
TASK
# ASSUME: ITP Tracker headers in row 1 include "ITP No", "Inspection Status", "Inspector" and "Remark"
# ASSUME: Closed is matched ignoring case and stray spaces; an Inspector holding only spaces counts as blank
# ASSUME: every row up to the last ITP No is checked, and the Remark of every row that is not a Closed point with no inspector is cleared
GOAL     Write No inspector in Remark for every Closed inspection point with a blank Inspector, clear the other remarks and return how many points have no inspector.
PROC     MarkClosedWithoutInspector
KIND     Function
MODULE   modItp
ARGS     ws As Worksheet
RET      Long
VAR      cItp As Long
VAR      cStatus As Long
VAR      cInspector As Long
VAR      cRemark As Long
VAR      marked As Long
PRE      COL($ws, "ITP No") > 0; COL($ws, "Inspection Status") > 0; COL($ws, "Inspector") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Remark"]
OP       LET(cItp, COL($ws, "ITP No"))
OP       LET(cStatus, COL($ws, "Inspection Status"))
OP       LET(cInspector, COL($ws, "Inspector"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cItp) var=r
OP         IF SAME(CELL($ws, r, cStatus), "Closed")
OP           IF IS_EMPTY(CELL($ws, r, cInspector))
OP             SET_VALUE(CELL($ws, r, cRemark), "No inspector")
OP             LET(marked, marked + 1)
OP           ELSE
OP             CLEAR(CELL($ws, r, cRemark), what=contents)
OP         ELSE
OP           CLEAR(CELL($ws, r, cRemark), what=contents)
OP       RETURN(marked)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  itp_01.xlsx sheet=Itp cols=A:ITP No,B:Point,C:Inspection Status,D:Inspector,E:Remark rows=10 data=A:seq(1,1);B:text(8);C:values("Closed"x3,"  closed "x1,"CLOSED"x1,"Open"x2,"Not closed"x1,""x1,"closed"x1)+boundary("Closed","  closed ","CLOSED","Closed","Open","Open","Not closed","","closed","Closed");D:values(""x6," "x1,"A. Lee"x1,"B. Ray"x1,"C. Ray"x1)+boundary("",""," ","A. Lee","","B. Ray","","","C. Ray","");E:values("old"x2,"No inspector"x2,""x6)+boundary("","old","","No inspector","No inspector","old","","","","")
FIXTURE  itp_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=itp_01.xlsx call=MarkClosedWithoutInspector(Itp) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(Itp![Remark], "No inspector") = 4 AND COUNTIF(Itp![Remark], "old") = 0 AND COUNTA(Itp![Remark]) = 5
TEST     fixture=itp_01.xlsx call=MarkClosedWithoutInspector(Itp) assert=VALUE(Itp![Remark]2) = "No inspector" AND VALUE(Itp![Remark]3) = "No inspector" AND VALUE(Itp![Remark]4) = "No inspector" AND VALUE(Itp![Remark]5) = "" AND VALUE(Itp![Remark]6) = "" AND VALUE(Itp![Remark]7) = "" AND VALUE(Itp![Remark]8) = "" AND VALUE(Itp![Remark]9) = "" AND VALUE(Itp![Remark]10) = "" AND VALUE(Itp![Remark]11) = "No inspector"
TEST     fixture=itp_01.xlsx call=MarkClosedWithoutInspector(Itp) assert=COUNTA(Itp![ITP No]) = PRE_COUNTA(Itp![ITP No]) AND COUNTA(Itp![Inspector]) = PRE_COUNTA(Itp![Inspector]) AND COUNTA(Itp![Inspection Status]) = PRE_COUNTA(Itp![Inspection Status])
TEST     fixture=itp_01.xlsx call=MarkClosedWithoutInspector(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
