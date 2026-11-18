VCL 0.6
TASK
# ASSUME: GRN headers in row 1 include "GRN No", "Received Date", "Inspection" and "Remark"
# ASSUME: more than 7 days ago means Received Date is before TODAY - 7; a line received exactly 7 days ago is not pending
# ASSUME: Inspection counts as blank when it is empty or only spaces; a Received Date that is not a real date is never pending
# ASSUME: the Remark is cleared on every line that is not pending; the last row is found from GRN No
GOAL     Write Inspection pending in Remark for each GRN line received over 7 days ago with no Inspection, clear the Remark on the others and return how many are pending.
PROC     MarkPendingInspections
KIND     Function
MODULE   modGrn
ARGS     ws As Worksheet
RET      Long
VAR      cGrn As Long
VAR      cDate As Long
VAR      cInsp As Long
VAR      cRemark As Long
VAR      pend As Long
VAR      marked As Long
PRE      COL($ws, "GRN No") > 0; COL($ws, "Received Date") > 0; COL($ws, "Inspection") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Remark"]
OP       LET(cGrn, COL($ws, "GRN No"))
OP       LET(cDate, COL($ws, "Received Date"))
OP       LET(cInsp, COL($ws, "Inspection"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cGrn) var=r
OP         LET(pend, 0)
OP         IF IS_DATE(CELL($ws, r, cDate))
OP           IF CELL($ws, r, cDate) < TODAY - 7
OP             IF IS_EMPTY(CELL($ws, r, cInsp)) THEN LET(pend, 1)
OP         IF pend = 1
OP           SET_VALUE(CELL($ws, r, cRemark), "Inspection pending")
OP           LET(marked, marked + 1)
OP         ELSE
OP           CLEAR(CELL($ws, r, cRemark), what=contents)
OP       RETURN(marked)
ERR      standard
FIXTURE  grn_01.xlsx sheet=Grn cols=A:GRN No,B:Part No,C:Received Date,D:Received Qty,E:Inspection,F:Remark rows=10 data=A:seq(1,1);B:text(6);C:reldates(-30,0)+boundary(-8,-7,-6,-8,"",-30,-9,0,-20,-1);D:int(1,50);E:values("Pass"x3,""x6," "x1)+boundary("","","","Pass","",""," ","Pass","Pass","");F:values("old"x7,"Inspection pending"x3)+boundary("old","Inspection pending","old","Inspection pending","old","old","old","old","Inspection pending","old")
FIXTURE  grn_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=grn_01.xlsx call=MarkPendingInspections(Grn) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Grn![Remark], "Inspection pending") = 3 AND COUNTIF(Grn![Remark], "old") = 0
TEST     fixture=grn_01.xlsx call=MarkPendingInspections(Grn) assert=VALUE(Grn![Remark]2) = "Inspection pending" AND VALUE(Grn![Remark]3) = "" AND VALUE(Grn![Remark]4) = "" AND VALUE(Grn![Remark]5) = "" AND VALUE(Grn![Remark]6) = "" AND VALUE(Grn![Remark]7) = "Inspection pending" AND VALUE(Grn![Remark]8) = "Inspection pending"
TEST     fixture=grn_01.xlsx call=MarkPendingInspections(Grn) assert=COUNTA(Grn![GRN No]) = PRE_COUNTA(Grn![GRN No]) AND COUNTA(Grn![Received Date]) = PRE_COUNTA(Grn![Received Date]) AND COUNTA(Grn![Inspection]) = PRE_COUNTA(Grn![Inspection]) AND SUM(Grn![Received Qty]) = PRE_SUM(Grn![Received Qty])
TEST     fixture=grn_01.xlsx call=MarkPendingInspections(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
