VCL 0.6
TASK
# ASSUME: ITP Tracker headers (row 1) include "ITP No", "Inspection Status" and "Planned Date"
# ASSUME: a Remark column is added after the last column when the sheet has none
# ASSUME: not Done means any status other than Done, so a blank status counts as not Done
# ASSUME: more than 7 days ago means a Planned Date before today minus 7; a blank or non-date Planned Date is never overdue
# ASSUME: every Remark on a row that is not overdue is cleared, whatever it said
# ASSUME: the real sheet is "ITP Tracker"; the TESTs use ITP
GOAL     Write Overdue in the Remark column of every inspection point that is not Done and was planned more than 7 days ago, clear the Remark on the rest and return how many are overdue.
PROC     MarkOverdueItpPoints
KIND     Function
MODULE   modItp
ARGS     ws As Worksheet
RET      Long
VAR      cItp As Long
VAR      cStat As Long
VAR      cDate As Long
VAR      cRem As Long
VAR      n As Long
PRE      COL($ws, "ITP No") > 0; COL($ws, "Inspection Status") > 0; COL($ws, "Planned Date") > 0
MODIFIES $ws
OP       LET(cItp, COL($ws, "ITP No"))
OP       LET(cStat, COL($ws, "Inspection Status"))
OP       LET(cDate, COL($ws, "Planned Date"))
OP       IF COL($ws, "Remark") = 0 THEN SET_VALUE(CELL($ws, 1, LASTCOL($ws, 1) + 1), "Remark")
OP       LET(cRem, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cItp) var=r
OP         CLEAR(CELL($ws, r, cRem), what=contents)
OP         IF NOT SAME(CELL($ws, r, cStat), "Done")
OP           IF IS_DATE(CELL($ws, r, cDate))
OP             IF CELL($ws, r, cDate) < TODAY - 7
OP               SET_VALUE(CELL($ws, r, cRem), "Overdue")
OP               LET(n, n + 1)
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  itp_01.xlsx sheet=ITP cols=A:ITP No,B:Point,C:Inspection Status,D:Planned Date,E:Remark rows=12 data=A:seq(1001,1);B:text(8);C:values("Pending"x6,"Done"x2,"DONE"x1,""x1,"Hold"x1,"Active"x1)+boundary("Pending","Pending","Pending","Done","DONE","","Hold","Pending","Pending","Done","Active","Pending");D:reldates(-100,20)+boundary(-8,-7,-6,-30,-30,-20,-40,"",15,-8,-100,-1);E:values("Overdue"x3,"note"x3,""x6)+boundary("","Overdue","note","Overdue","note","","Overdue","","","note","","")
FIXTURE  itp_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  itp_02.xlsx sheet=ITP cols=A:ITP No,B:Point,C:Inspection Status,D:Planned Date rows=7 data=A:seq(2001,1);B:text(8);C:values("Open"x5,"Done"x2)+boundary("Open","Open","Done","Open","Open","Done","Open");D:reldates(-30,10)+boundary("",-9,-30,-7,5,-20,-8)
TEST     fixture=itp_01.xlsx call=MarkOverdueItpPoints(ITP) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(ITP![Remark], "Overdue") = 4 AND COUNTA(ITP![Remark]) = 5 AND COUNTIF(ITP![Remark], "note") = 0 AND COUNTIFS(ITP![Remark], "Overdue", ITP![Inspection Status], "Hold") = 1
TEST     fixture=itp_01.xlsx call=MarkOverdueItpPoints(ITP) assert=RESULT = 4 AND COUNTA(ITP![ITP No]) = PRE_COUNTA(ITP![ITP No]) AND COUNTA(ITP![Inspection Status]) = PRE_COUNTA(ITP![Inspection Status]) AND COUNTA(ITP![Planned Date]) = PRE_COUNTA(ITP![Planned Date]) AND VALUE(ITP![Remark]2) = "Overdue" AND VALUE(ITP![Remark]7) = "Overdue" AND VALUE(ITP![Remark]3) = ""
TEST     fixture=itp_02.xlsx call=MarkOverdueItpPoints(ITP) assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(ITP![Remark]1) = "Remark" AND VALUE(ITP!E1) = "Remark" AND COUNTIF(ITP![Remark], "Overdue") = 2
TEST     fixture=itp_01.xlsx call=MarkOverdueItpPoints(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
