VCL 0.6
TASK
# ASSUME: Welder Quals headers in row 1 include "Welder ID" and "Remark"
# ASSUME: the ID is matched ignoring case and stray spaces; a blank ID given to the macro finds nothing and returns 0
# ASSUME: only the first matching row is changed, and its Remark is overwritten
GOAL     Find the first Welder Quals row for the given welder ID, write Renewal requested in its Remark and return that row, or 0 if not found.
PROC     RequestRenewal
KIND     Function
MODULE   modQuals
ARGS     ws As Worksheet, welderId As String
RET      Long
VAR      cId As Long
VAR      cRemark As Long
VAR      found As Long
PRE      COL($ws, "Welder ID") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Remark"]
OP       EXIT_IF IS_EMPTY($welderId)
OP       LET(cId, COL($ws, "Welder ID"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cId) var=r
OP         SKIP_IF NOT SAME(CELL($ws, r, cId), $welderId)
OP         SET_VALUE(CELL($ws, r, cRemark), "Renewal requested")
OP         LET(found, r)
OP         EXIT_LOOP_IF found > 0
OP       RETURN(found)
PERF     screen_off
ERR      standard
FIXTURE  quals_01.xlsx sheet=Quals cols=A:Welder ID,B:Process,C:Expiry Date,D:Remark rows=8 data=A:values("W-100"x1,"W-101"x2,"  w-102 "x1,"W-102"x1,""x1,"W-103"x1,"W-104"x1)+boundary("W-100","W-101","  w-102 ","W-101","","W-102","W-103");B:values("GTAW"x4,"SMAW"x4);C:dates(2026-01-01,2027-12-31);D:values("old"x8)
FIXTURE  quals_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=quals_01.xlsx call=RequestRenewal(Quals, "W-101") runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Quals![Remark], "Renewal requested") = 1 AND VALUE(Quals![Remark]3) = "Renewal requested" AND VALUE(Quals![Remark]5) = "old"
TEST     fixture=quals_01.xlsx call=RequestRenewal(Quals, "w-102 ") assert=RESULT = 4 AND COUNTIF(Quals![Remark], "Renewal requested") = 1 AND VALUE(Quals![Remark]4) = "Renewal requested" AND COUNTA(Quals![Welder ID]) = PRE_COUNTA(Quals![Welder ID])
TEST     fixture=quals_01.xlsx call=RequestRenewal(Quals, "W-999") assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(Quals![Remark], "Renewal requested") = 0
TEST     fixture=quals_01.xlsx call=RequestRenewal(Quals, "") assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(Quals![Remark], "Renewal requested") = 0
TEST     fixture=quals_01.xlsx call=RequestRenewal(Quals, "W-100") assert=RESULT = 2 AND VALUE(Quals![Remark]2) = "Renewal requested"
TEST     fixture=quals_01.xlsx call=RequestRenewal(Notes, "W-101") assert=RESULT = -1 AND ERROR_COUNT = 1
