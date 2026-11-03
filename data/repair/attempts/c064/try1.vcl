VCL 0.6
TASK
# ASSUME: NDE Report headers in row 1 include "Weld Seam No", "Method", "Result" and "Report Date"
# ASSUME: the "Rejected Welds" sheet already exists; it is rebuilt on every run with the four NDE Report headings in that order
# ASSUME: a row counts as rejected when Result is Reject, matched ignoring case and extra spaces; Rejected does not match
# ASSUME: Report Date holds real dates
# ASSUME: the real sheets are "NDE Report" and "Rejected Welds"; the TESTs use NdeReport and RejectedWelds
GOAL     Rebuild the Rejected Welds sheet with every NDE Report row whose Result is Reject, newest Report Date first, and return how many rows were listed.
PROC     ListRejectedWelds
KIND     Function
MODULE   modNde
ARGS     nde As Worksheet, out As Worksheet
RET      Long
VAR      cSeam As Long
VAR      cMethod As Long
VAR      cResult As Long
VAR      cDate As Long
VAR      listed As Long
PRE      COL($nde, "Weld Seam No") > 0; COL($nde, "Method") > 0; COL($nde, "Result") > 0; COL($nde, "Report Date") > 0
MODIFIES $out
OP       LET(cSeam, COL($nde, "Weld Seam No"))
OP       LET(cMethod, COL($nde, "Method"))
OP       LET(cResult, COL($nde, "Result"))
OP       LET(cDate, COL($nde, "Report Date"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, CELL($nde, 1, cSeam), CELL($nde, 1, cMethod), CELL($nde, 1, cResult), CELL($nde, 1, cDate))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($nde, cResult) var=r
OP         IF SAME(CELL($nde, r, cResult), "Reject")
OP           LET(listed, listed + 1)
OP           WRITE_ROW($out, listed + 1, CELL($nde, r, cSeam), CELL($nde, r, cMethod), CELL($nde, r, cResult), CELL($nde, r, cDate))
OP       IF listed > 1
OP         SORT(USED_RANGE($out), key=4, order=desc, header=yes)
OP       RETURN(listed)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  nde_01.xlsx sheet=NdeReport cols=A:Weld Seam No,B:Method,C:Result,D:Report Date rows=10 data=A:seq(1,1);B:values("RT"x4,"UT"x3,"MT"x2,"PT"x1);C:values("Accept"x3,"Reject"x3," reject "x1,"Rejected"x1,"REJECT"x1,""x1)+boundary("Accept","Reject"," reject ","Rejected","REJECT","","Accept","Reject","Accept","Reject");D:dates(2026-01-01,2026-09-30)+boundary(2026-03-10,2026-04-02,2026-06-15,2026-07-01,2026-05-20,2026-08-01,2026-02-02,2026-01-12,2026-03-03,2026-09-09)
FIXTURE  nde_01.xlsx sheet=RejectedWelds cols=A:Weld Seam No,B:Method,C:Result,D:Report Date rows=0 stale=3
FIXTURE  nde_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  nde_02.xlsx sheet=NdeReport cols=A:Weld Seam No,B:Method,C:Result,D:Report Date rows=4 data=A:seq(1,1);B:values("RT"x2,"UT"x2);C:values("Accept"x3,"Rejected"x1)+boundary("Accept","Rejected","Accept","Accept");D:dates(2026-01-01,2026-09-30)
FIXTURE  nde_02.xlsx sheet=RejectedWelds cols=A:Weld Seam No,B:Method,C:Result,D:Report Date rows=0 stale=2
FIXTURE  nde_03.xlsx sheet=NdeReport cols=A:Weld Seam No,B:Method,C:Result,D:Report Date rows=3 data=A:seq(1,1);B:values("RT"x1,"UT"x2);C:values("Accept"x2,"Reject"x1)+boundary("Accept","Reject","Accept");D:dates(2026-01-01,2026-09-30)+boundary(2026-04-04,2026-05-05,2026-06-06)
FIXTURE  nde_03.xlsx sheet=RejectedWelds cols=A:Weld Seam No,B:Method,C:Result,D:Report Date rows=0 stale=2
TEST     fixture=nde_01.xlsx call=ListRejectedWelds(NdeReport, RejectedWelds) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND ROWS_USED(RejectedWelds) = 6 AND COUNTA(RejectedWelds![Weld Seam No]) = 6 AND COUNTIF(RejectedWelds![Weld Seam No], "stale") = 0
TEST     fixture=nde_01.xlsx call=ListRejectedWelds(NdeReport, RejectedWelds) assert=VALUE(RejectedWelds![Weld Seam No]1) = "Weld Seam No" AND VALUE(RejectedWelds![Method]1) = "Method" AND VALUE(RejectedWelds![Result]1) = "Result" AND VALUE(RejectedWelds![Report Date]1) = "Report Date" AND VALUE(RejectedWelds![Weld Seam No]2) = 10 AND VALUE(RejectedWelds![Weld Seam No]3) = 3 AND VALUE(RejectedWelds![Weld Seam No]4) = 5 AND VALUE(RejectedWelds![Weld Seam No]5) = 2 AND VALUE(RejectedWelds![Weld Seam No]6) = 8
TEST     fixture=nde_01.xlsx call=ListRejectedWelds(NdeReport, RejectedWelds) assert=VALUE(RejectedWelds![Result]3) = " reject " AND VALUE(RejectedWelds![Result]4) = "REJECT" AND COUNTIF(RejectedWelds![Result], "Rejected") = 0 AND COUNTA(RejectedWelds![Method]) = 6 AND COUNTA(RejectedWelds![Report Date]) = 6
TEST     fixture=nde_01.xlsx call=ListRejectedWelds(NdeReport, RejectedWelds) assert=ERROR_COUNT = 0 AND COUNTA(NdeReport![Weld Seam No]) = PRE_COUNTA(NdeReport![Weld Seam No]) AND COUNTA(NdeReport![Result]) = PRE_COUNTA(NdeReport![Result]) AND COUNTA(NdeReport![Report Date]) = PRE_COUNTA(NdeReport![Report Date])
TEST     fixture=nde_02.xlsx call=ListRejectedWelds(NdeReport, RejectedWelds) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND ROWS_USED(RejectedWelds) = 1 AND COUNTA(RejectedWelds![Weld Seam No]) = 1 AND COUNTIF(RejectedWelds![Weld Seam No], "stale") = 0
TEST     fixture=nde_03.xlsx call=ListRejectedWelds(NdeReport, RejectedWelds) assert=RESULT = 1 AND ERROR_COUNT = 0 AND ROWS_USED(RejectedWelds) = 2 AND VALUE(RejectedWelds![Weld Seam No]2) = 2 AND VALUE(RejectedWelds![Report Date]1) = "Report Date"
TEST     fixture=nde_01.xlsx call=ListRejectedWelds(Notes, RejectedWelds) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(RejectedWelds![Weld Seam No]) = PRE_COUNTA(RejectedWelds![Weld Seam No])
