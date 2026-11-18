VCL 0.6
TASK
# ASSUME: ITP Tracker headers in row 1 include "ITP No", "Point", "Inspection Status" and "Planned Date"
# ASSUME: the real sheets are "ITP Tracker" and "Open Points"; Open Points exists and is rebuilt on every run
# ASSUME: a point is not Closed when its status is anything but Closed (case and spaces ignored), a blank status included; rows with a blank ITP No are skipped
# ASSUME: Planned Date holds real dates; a blank date sorts to the end
GOAL     Rebuild the Open Points sheet with every ITP point that is not Closed, earliest planned date first, and return how many were listed.
PROC     ListOpenPoints
KIND     Function
MODULE   modItp
ARGS     itp As Worksheet, out As Worksheet
RET      Long
VAR      cNo As Long
VAR      cPoint As Long
VAR      cStatus As Long
VAR      cDate As Long
VAR      outRow As Long
PRE      COL($itp, "ITP No") > 0; COL($itp, "Point") > 0; COL($itp, "Inspection Status") > 0; COL($itp, "Planned Date") > 0
MODIFIES $out
OP       LET(cNo, COL($itp, "ITP No"))
OP       LET(cPoint, COL($itp, "Point"))
OP       LET(cStatus, COL($itp, "Inspection Status"))
OP       LET(cDate, COL($itp, "Planned Date"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "ITP No", "Point", "Inspection Status", "Planned Date")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($itp, cNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($itp, r, cNo))
OP         SKIP_IF SAME(CELL($itp, r, cStatus), "Closed")
OP         WRITE_ROW($out, outRow, CELL($itp, r, cNo), CELL($itp, r, cPoint), CELL($itp, r, cStatus), CELL($itp, r, cDate))
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=4, order=asc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off
ERR      standard
FIXTURE  itp_01.xlsx sheet=Itp cols=A:ITP No,B:Point,C:Inspection Status,D:Planned Date rows=8 data=A:seq(1,1)+boundary(1,2,"",4);B:text(10);C:values("Open"x2,"Closed"x2,"  closed "x1,"Pending"x1,"Hold"x1,""x1)+boundary("Open","Closed","Open","  closed ","Hold","Pending","","Closed");D:dates(2026-01-01,2026-12-31)+boundary(2026-05-01,2026-03-01,2026-02-01,"",2026-04-01,"",2026-06-01)
FIXTURE  itp_01.xlsx sheet=OpenPts cols=A:ITP No,B:Point,C:Inspection Status,D:Planned Date rows=0 stale=2
TEST     fixture=itp_01.xlsx call=ListOpenPoints(Itp, OpenPts) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(OpenPts![ITP No]) = 5 AND COUNTA(OpenPts![Planned Date]) = 4 AND COUNTIF(OpenPts![Inspection Status], "stale") = 0 AND COUNTIF(OpenPts![Inspection Status], "Closed") = 0
TEST     fixture=itp_01.xlsx call=ListOpenPoints(Itp, OpenPts) assert=COUNTA(Itp![ITP No]) = PRE_COUNTA(Itp![ITP No]) AND COUNTA(Itp![Planned Date]) = PRE_COUNTA(Itp![Planned Date]) AND VALUE(OpenPts![Planned Date]2) < VALUE(OpenPts![Planned Date]3) AND VALUE(OpenPts![Planned Date]3) < VALUE(OpenPts![Planned Date]4) AND VALUE(OpenPts![Inspection Status]5) = "Pending" AND VALUE(OpenPts![ITP No]2) = 6
TEST     fixture=itp_01.xlsx call=ListOpenPoints(OpenPts, Itp) assert=RESULT = -1 AND ERROR_COUNT = 1
