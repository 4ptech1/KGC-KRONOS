VCL 0.6
TASK
# ASSUME: the sheets are "ITP Tracker" and "Open Points"; Open Points exists and is rebuilt on every run
# ASSUME: ITP Tracker headers in row 1 include "ITP No", "Point", "Inspection Status" and "Planned Date"
# ASSUME: a point is Closed when its Inspection Status is Closed, ignoring case and spaces; a blank status is not Closed
# ASSUME: rows with a blank ITP No are skipped
# ASSUME: Planned Date holds real dates; a blank or non-date Planned Date is treated as no date and left blank in the list
GOAL     List every ITP point that is not Closed on the Open Points sheet, earliest planned date first and undated points last.
PROC     ListOpenPoints
KIND     Function
MODULE   modItp
ARGS     itp As Worksheet, out As Worksheet
RET      Long
VAR      cNo As Long
VAR      cPoint As Long
VAR      cStat As Long
VAR      cDate As Long
VAR      outRow As Long
PRE      COL($itp, "ITP No") > 0; COL($itp, "Point") > 0; COL($itp, "Inspection Status") > 0; COL($itp, "Planned Date") > 0
MODIFIES $out
OP       LET(cNo, COL($itp, "ITP No"))
OP       LET(cPoint, COL($itp, "Point"))
OP       LET(cStat, COL($itp, "Inspection Status"))
OP       LET(cDate, COL($itp, "Planned Date"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "ITP No", "Point", "Inspection Status", "Planned Date")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($itp, cNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($itp, r, cNo))
OP         SKIP_IF SAME(CELL($itp, r, cStat), "Closed")
OP         IF IS_DATE(CELL($itp, r, cDate))
OP           WRITE_ROW($out, outRow, CELL($itp, r, cNo), CELL($itp, r, cPoint), CELL($itp, r, cStat), CELL($itp, r, cDate))
OP         ELSE
OP           WRITE_ROW($out, outRow, CELL($itp, r, cNo), CELL($itp, r, cPoint), CELL($itp, r, cStat))
OP         LET(outRow, outRow + 1)
OP       IF outRow > 3
OP         SORT(USED_RANGE($out), key=4, order=asc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  itp_01.xlsx sheet=Itp cols=A:ITP No,B:Point,C:Inspection Status,D:Planned Date rows=12 data=A:seq(1,1)+boundary("");B:text(8);C:values("Open"x4,"closed "x1,"Closed"x2,"Hold"x2,""x1,"Witness"x2)+boundary("Open","closed ","Closed","Hold","","Witness","Open","Open","Closed","Hold","Witness","Open");D:dates(2026-01-01,2026-12-31)+boundary(2026-06-01,"",2026-03-01,2026-06-01,"",2026-03-01,"",2026-01-15,"",2026-06-01,2026-09-09,2026-02-02)
FIXTURE  itp_01.xlsx sheet=Out cols=A:ITP No,B:Point,C:Inspection Status,D:Planned Date rows=0 stale=3
TEST     fixture=itp_01.xlsx call=ListOpenPoints(Itp, Out) runs=2 assert=RESULT = 8 AND ERROR_COUNT = 0 AND COUNTA(Out![ITP No]) = 9 AND COUNTIF(Out![ITP No], "stale") = 0 AND COUNTIF(Out![Inspection Status], "Closed") = 0 AND COUNTIF(Out![Inspection Status], "closed ") = 0 AND COUNTIF(Out![Inspection Status], "Hold") = 2
TEST     fixture=itp_01.xlsx call=ListOpenPoints(Itp, Out) assert=VALUE(Out![Planned Date]2) <= VALUE(Out![Planned Date]3) AND VALUE(Out![Planned Date]3) <= VALUE(Out![Planned Date]4) AND VALUE(Out![Planned Date]6) <= VALUE(Out![Planned Date]7) AND VALUE(Out![Planned Date]7) < VALUE(Out![Planned Date]4) + 300 AND COUNTA(Out![Planned Date]) = 7 AND COUNTA(Out![Point]) = 9 AND COUNTIF(Out![Inspection Status], "Witness") = 2 AND COUNTA(Itp![ITP No]) = PRE_COUNTA(Itp![ITP No]) AND COUNTA(Itp![Planned Date]) = PRE_COUNTA(Itp![Planned Date])
TEST     fixture=itp_01.xlsx call=ListOpenPoints(Out, Itp) assert=RESULT = -1 AND ERROR_COUNT = 1
