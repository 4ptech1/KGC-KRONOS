VCL 0.6
TASK
# ASSUME: GRN headers in row 1 include "GRN No", "Received Date", "Inspector" and "Remark"; the Remark column already exists
# ASSUME: the Remark column holds only this tool's remarks, so every Remark cell below the heading is cleared on each run
# ASSUME: old highlights means any fill on a GRN No cell; fills on other columns are not touched
# ASSUME: an Inspector cell counts as blank when it is empty or holds only spaces
# ASSUME: Received Date holds real dates; a row with no date, or a date typed as text, gets no remark; today itself is not a future date
GOAL     Highlight the GRN No cell yellow when the inspector is blank, write Future date in Remark when the received date is after today, and return how many GRN No cells are yellow.
PROC     MarkGrnIssues
KIND     Function
MODULE   modGrnCheck
ARGS     ws As Worksheet
RET      Long
VAR      cGrn As Long
VAR      cDate As Long
VAR      cInsp As Long
VAR      cRemark As Long
VAR      marked As Long
PRE      COL($ws, "GRN No") > 0; COL($ws, "Received Date") > 0; COL($ws, "Inspector") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Remark"]
MODIFIES $ws fills
OP       LET(cGrn, COL($ws, "GRN No"))
OP       LET(cDate, COL($ws, "Received Date"))
OP       LET(cInsp, COL($ws, "Inspector"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cGrn) var=r
OP         FORMAT(CELL($ws, r, cGrn), fill=none)
OP         CLEAR(CELL($ws, r, cRemark), what=contents)
OP         IF IS_EMPTY(CELL($ws, r, cInsp))
OP           FORMAT(CELL($ws, r, cGrn), fill=yellow)
OP           LET(marked, marked + 1)
OP         IF IS_DATE(CELL($ws, r, cDate))
OP           IF CELL($ws, r, cDate) > TODAY
OP             SET_VALUE(CELL($ws, r, cRemark), "Future date")
OP       RETURN(marked)
PERF     screen_off
ERR      standard
FIXTURE  grn_01.xlsx sheet=GRN cols=A:GRN No,B:Part No,C:Received Qty,D:Received Date,E:Inspector,F:Remark rows=9 stale=3 data=A:seq(1,1);B:text(6);C:int(1,100);D:reldates(-30,30)+boundary(1,0,-1,"",5,-5,30,-30,-10);E:values(""x3,"  "x1,"Kim"x5)+boundary("","  ","Kim","Kim","Kim","","Kim","Kim","");F:values("old"x9)
FIXTURE  grn_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=grn_01.xlsx call=MarkGrnIssues(GRN) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND FILL_COUNT(GRN![GRN No], #FFFF00) = 4 AND FILL_COUNT(GRN![GRN No], #FF0000) = 0 AND COUNTIF(GRN![Remark], "Future date") = 3 AND COUNTIF(GRN![Remark], "old") = 0 AND COUNTA(GRN![Remark]) = 4
TEST     fixture=grn_01.xlsx call=MarkGrnIssues(GRN) assert=COUNTA(GRN![GRN No]) = PRE_COUNTA(GRN![GRN No]) AND COUNTA(GRN![Part No]) = PRE_COUNTA(GRN![Part No]) AND SUM(GRN![Received Qty]) = PRE_SUM(GRN![Received Qty]) AND COUNTA(GRN![Received Date]) = PRE_COUNTA(GRN![Received Date]) AND COUNTA(GRN![Inspector]) = PRE_COUNTA(GRN![Inspector]) AND FILL_COUNT(GRN![Part No], #FF0000) = 3 AND FILL_COUNT(GRN![Part No], #FFFF00) = 0 AND VALUE(GRN![Remark]2) = "Future date" AND VALUE(GRN![Remark]7) = "Future date"
TEST     fixture=grn_01.xlsx call=MarkGrnIssues(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(GRN![Remark]) = PRE_COUNTA(GRN![Remark])
