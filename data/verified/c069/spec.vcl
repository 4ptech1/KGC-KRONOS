VCL 0.6
TASK
# ASSUME: NDE Report headers in row 1 include "Weld Seam No", "Method" and "Result"; only Result is used; the real sheet is "NDE Report", the TESTs use NdeReport
# ASSUME: a Reject row has a Result of Reject, ignoring case and spaces at the ends; every other row, a blank Result included, is hidden
# ASSUME: the heading row stays visible, no row is deleted, and the number of visible rows below the headings is the Function's result
GOAL     Hide every row of the NDE report whose Result is not Reject so that only the rejected welds show, and report how many rows are visible below the headings.
PROC     ShowOnlyRejectedWelds
KIND     Function
MODULE   modNde
ARGS     ws As Worksheet
RET      Long
VAR      cResult As Long
VAR      shown As Long
PRE      COL($ws, "Result") > 0
MODIFIES $ws
OP       LET(cResult, COL($ws, "Result"))
OP       FILTER(USED_RANGE($ws), col=cResult, crit="Reject")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cResult) var=r
OP         IF SAME(CELL($ws, r, cResult), "Reject") THEN LET(shown, shown + 1)
OP       RETURN(shown)
PERF     screen_off
ERR      standard
FIXTURE  nde_report_01.xlsx sheet=NdeReport cols=A:Weld Seam No,B:Method,C:Result rows=15 data=A:seq(101,1);B:values("RT"x5,"UT"x4,"MT"x3,"PT"x3);C:values("Accept"x5,"Reject"x3,"reject "x1," REJECT"x1,"Repair"x2,""x2,"Rejected"x1)+boundary("Reject","Accept","Repair","reject ","","Accept","Rejected"," REJECT","Accept","Reject","Repair","","Accept","Accept","Reject")
FIXTURE  nde_report_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=nde_report_01.xlsx call=ShowOnlyRejectedWelds(NdeReport) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND VISIBLE_COUNT(NdeReport![Weld Seam No]) = 6 AND VISIBLE_COUNT(NdeReport![Result]) = 6
TEST     fixture=nde_report_01.xlsx call=ShowOnlyRejectedWelds(NdeReport) assert=VISIBLE_COUNT(NdeReport![Weld Seam No]2) = 1 AND VISIBLE_COUNT(NdeReport![Weld Seam No]5) = 1 AND VISIBLE_COUNT(NdeReport![Weld Seam No]9) = 1 AND VISIBLE_COUNT(NdeReport![Weld Seam No]11) = 1 AND VISIBLE_COUNT(NdeReport![Weld Seam No]16) = 1 AND VISIBLE_COUNT(NdeReport![Weld Seam No]3) = 0 AND VISIBLE_COUNT(NdeReport![Weld Seam No]4) = 0 AND VISIBLE_COUNT(NdeReport![Weld Seam No]6) = 0 AND VISIBLE_COUNT(NdeReport![Weld Seam No]8) = 0 AND VISIBLE_COUNT(NdeReport![Weld Seam No]15) = 0 AND COUNTA(NdeReport![Weld Seam No]) = PRE_COUNTA(NdeReport![Weld Seam No]) AND COUNTA(NdeReport![Result]) = PRE_COUNTA(NdeReport![Result]) AND COUNTIF(NdeReport![Result], "Accept") = PRE_COUNTIF(NdeReport![Result], "Accept")
TEST     fixture=nde_report_01.xlsx call=ShowOnlyRejectedWelds(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
