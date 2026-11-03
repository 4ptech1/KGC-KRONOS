VCL 0.6
TASK
# ASSUME: ITP Tracker headers in row 1 include "ITP No", "Inspection Status" and "Planned Date"; "Flag" is added after the last heading when it is missing
# ASSUME: every point row has an ITP No, and the last row is found from it
# ASSUME: a point is overdue when its Inspection Status is not Closed (a blank status counts as not Closed; case and extra spaces ignored) and its Planned Date is a real date before today
# ASSUME: a point whose Planned Date is blank or not a real date is never flagged; every point that is not overdue has its Flag cleared
GOAL     Write Overdue in the Flag column of every ITP point that is not Closed and planned before today, clear the Flag on all other points, and return how many are overdue.
PROC     FlagOverduePoints
KIND     Function
MODULE   modItp
ARGS     ws As Worksheet
RET      Long
VAR      cNo As Long
VAR      cStatus As Long
VAR      cDate As Long
VAR      cFlag As Long
VAR      n As Long
PRE      COL($ws, "ITP No") > 0; COL($ws, "Inspection Status") > 0; COL($ws, "Planned Date") > 0
MODIFIES $ws
OP       LET(cNo, COL($ws, "ITP No"))
OP       LET(cStatus, COL($ws, "Inspection Status"))
OP       LET(cDate, COL($ws, "Planned Date"))
OP       IF COL($ws, "Flag") = 0
OP         SET_VALUE(CELL($ws, 1, LASTCOL($ws, 1) + 1), "Flag")
OP       LET(cFlag, COL($ws, "Flag"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cNo) var=r
OP         CLEAR(CELL($ws, r, cFlag), what=contents)
OP         IF NOT SAME(CELL($ws, r, cStatus), "Closed")
OP           IF IS_DATE(CELL($ws, r, cDate))
OP             IF CELL($ws, r, cDate) < TODAY
OP               SET_VALUE(CELL($ws, r, cFlag), "Overdue")
OP               LET(n, n + 1)
OP       RETURN(n)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  itp_01.xlsx sheet=ITP cols=A:ITP No,B:Point,C:Inspection Status,D:Planned Date,E:Flag rows=12 data=A:seq(1,1);B:text(10);C:values("Open"x6,"Closed"x2," closed "x1,""x1,"Hold"x2)+boundary("Open","Open","Open","Closed"," closed ","","Hold","Open","Hold","Closed","Open","Open");D:reldates(-30,30)+boundary(-1,0,1,-5,-5,-2,"",-10,-3,3,-30,30);E:values("Overdue"x5,"Review"x3,""x4)+boundary("","Overdue","Review","Overdue","","Review","Overdue","Overdue","","Review","","Overdue")
FIXTURE  itp_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  itp_02.xlsx sheet=ITP cols=A:ITP No,B:Point,C:Inspection Status,D:Planned Date rows=6 data=A:seq(1,1);B:text(10);C:values("Open"x4,"Closed"x1,"Hold"x1)+boundary("Open","Open","Hold","Open","Open","Closed");D:reldates(-20,20)+boundary("",-1,-4,0,2,-9)
TEST     fixture=itp_01.xlsx call=FlagOverduePoints(ITP) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTIF(ITP![Flag], "Overdue") = 5 AND COUNTA(ITP![Flag]) = 6 AND VALUE(ITP![Flag]2) = "Overdue" AND VALUE(ITP![Flag]9) = "Overdue"
TEST     fixture=itp_01.xlsx call=FlagOverduePoints(ITP) assert=RESULT = 5 AND COUNTA(ITP![ITP No]) = PRE_COUNTA(ITP![ITP No]) AND COUNTA(ITP![Inspection Status]) = PRE_COUNTA(ITP![Inspection Status]) AND COUNTIF(ITP![Inspection Status], "Closed") = PRE_COUNTIF(ITP![Inspection Status], "Closed") AND COUNTA(ITP![Planned Date]) = PRE_COUNTA(ITP![Planned Date])
TEST     fixture=itp_02.xlsx call=FlagOverduePoints(ITP) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(ITP!E1) = "Flag" AND COUNTIF(ITP![Flag], "Overdue") = 2 AND COUNTA(ITP![Flag]) = 3 AND COUNTA(ITP!F:F) = 0
TEST     fixture=itp_01.xlsx call=FlagOverduePoints(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
