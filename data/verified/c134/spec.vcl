VCL 0.6
TASK
# ASSUME: Weld Log headers in row 1 include "Weld Seam No", "Visual", "NDE Required", "NDE Result" and "Remark"
# ASSUME: Accepted and Yes are matched ignoring case and extra spaces; NDE Result counts as blank when empty or only spaces
# ASSUME: rows are walked down to the last Weld Seam No; every row that is not pending has its Remark cleared
GOAL     Write NDE pending in Remark for every accepted weld that needs NDE and has no NDE result yet, clear Remark on the other welds and return how many are pending.
PROC     MarkNdePending
KIND     Function
MODULE   modWeldLog
ARGS     ws As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cVisual As Long
VAR      cReq As Long
VAR      cRes As Long
VAR      cRemark As Long
VAR      pending As Long
PRE      COL($ws, "Weld Seam No") > 0; COL($ws, "Visual") > 0; COL($ws, "NDE Required") > 0; COL($ws, "NDE Result") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Remark"]
OP       LET(cSeam, COL($ws, "Weld Seam No"))
OP       LET(cVisual, COL($ws, "Visual"))
OP       LET(cReq, COL($ws, "NDE Required"))
OP       LET(cRes, COL($ws, "NDE Result"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cSeam) var=r
OP         IF SAME(CELL($ws, r, cVisual), "Accepted") AND SAME(CELL($ws, r, cReq), "Yes") AND IS_EMPTY(CELL($ws, r, cRes))
OP           SET_VALUE(CELL($ws, r, cRemark), "NDE pending")
OP           LET(pending, pending + 1)
OP         ELSE
OP           CLEAR(CELL($ws, r, cRemark), what=contents)
OP       RETURN(pending)
PERF     screen_off
ERR      standard
FIXTURE  weld_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Visual,C:NDE Required,D:NDE Result,E:Remark rows=10 data=A:seq(1,1);B:values("Accepted"x5,"accepted "x1,"ACCEPTED"x1,"Rejected"x1,"Repair"x1,""x1)+boundary("Accepted","accepted ","ACCEPTED","Rejected","Accepted","Accepted","Repair","","Accepted","Accepted");C:values("Yes"x6,"YES"x1," yes "x1,"No"x1,""x1)+boundary("Yes","YES"," yes ","Yes","No","Yes","Yes","Yes","","Yes");D:values(""x7," "x1,"RT OK"x1,"UT OK"x1)+boundary(""," ","","","","RT OK","","","","UT OK");E:values("old"x10)
FIXTURE  weld_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=weld_01.xlsx call=MarkNdePending(WeldLog) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(WeldLog![Remark], "NDE pending") = 3 AND COUNTIF(WeldLog![Remark], "old") = 0 AND COUNTA(WeldLog![Remark]) = 4
TEST     fixture=weld_01.xlsx call=MarkNdePending(WeldLog) assert=COUNTA(WeldLog![Weld Seam No]) = PRE_COUNTA(WeldLog![Weld Seam No]) AND COUNTA(WeldLog![NDE Result]) = PRE_COUNTA(WeldLog![NDE Result]) AND VALUE(WeldLog![Remark]2) = "NDE pending" AND VALUE(WeldLog![Remark]3) = "NDE pending" AND VALUE(WeldLog![Remark]4) = "NDE pending" AND VALUE(WeldLog![Remark]5) = "" AND VALUE(WeldLog![Remark]6) = "" AND VALUE(WeldLog![Remark]7) = "" AND VALUE(WeldLog![Remark]8) = "" AND VALUE(WeldLog![Remark]9) = "" AND VALUE(WeldLog![Remark]10) = "" AND VALUE(WeldLog![Remark]11) = ""
TEST     fixture=weld_01.xlsx call=MarkNdePending(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
