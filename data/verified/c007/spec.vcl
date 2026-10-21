VCL 0.6
TASK
# ASSUME: ITP Tracker headers (row 1) include "Point Type" (Hold / Witness / Review) and "Inspection Status"
# ASSUME: a row is copied when Inspection Status is "Ready" and Point Type is "Hold" or "Witness" (case and spaces ignored)
# ASSUME: rows with a blank Inspection Status are ignored
# ASSUME: the Call Notice sheet is rebuilt every run with the ITP Tracker's header row
GOAL     Rebuild a Call Notice sheet with every ITP point of type Hold or Witness whose inspection status is Ready.
PROC     BuildCallNoticeList
KIND     Function
MODULE   modQuality
ARGS     itp As Worksheet, out As Worksheet
RET      Long
VAR      cType As Long
VAR      cStatus As Long
VAR      lastCol As Long
VAR      outRow As Long
VAR      pointType As String
VAR      status As String
PRE      COL($itp, "Point Type") > 0; COL($itp, "Inspection Status") > 0
MODIFIES $out
OP       LET(cType, COL($itp, "Point Type"))
OP       LET(cStatus, COL($itp, "Inspection Status"))
OP       LET(lastCol, LASTCOL($itp, 1))
OP       CLEAR(USED_RANGE($out), what=all)
OP       COPY_ROW($itp, 1, $out, 1, cols=lastCol)
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($itp, cStatus) var=r
OP         LET(status, UPPER(TRIM(CELL($itp, r, cStatus))))
OP         LET(pointType, UPPER(TRIM(CELL($itp, r, cType))))
OP         IF status = "READY"
OP           IF pointType = "HOLD" OR pointType = "WITNESS"
OP             COPY_ROW($itp, r, $out, outRow, cols=lastCol)
OP             LET(outRow, outRow + 1)
OP       RETURN(outRow - 2)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  itp_01.xlsx sheet=ITP cols=A:Activity,B:Point Type,C:Inspection Status,D:Planned Date rows=40 data=A:text(15);B:values("Hold"x12,"Witness"x14,"Review"x10,""x4)+boundary("Hold");C:values("Ready"x10,"Done"x18,"Pending"x8,""x4)+boundary("Ready");D:dates(2026-10-01,2026-12-31)+boundary("")
FIXTURE  itp_01.xlsx sheet=CallNotice cols=A:Activity rows=0 stale=10
TEST     fixture=itp_01.xlsx call=BuildCallNoticeList(ITP, CallNotice) runs=2 assert=RESULT = 8 AND RESULT = PRE_COUNTIFS(ITP![Inspection Status], "Ready", ITP![Point Type], "Hold") + PRE_COUNTIFS(ITP![Inspection Status], "Ready", ITP![Point Type], "Witness") AND COUNTA(CallNotice![Activity]) = 9 AND COUNTIF(CallNotice![Activity], "stale") = 0 AND COUNTA(CallNotice![Planned Date]) = 8 AND ERROR_COUNT = 0
TEST     fixture=itp_01.xlsx call=BuildCallNoticeList(ITP, CallNotice) assert=COUNTIF(CallNotice![Inspection Status], "Ready") = RESULT AND COUNTIF(CallNotice![Point Type], "Review") = 0
TEST     fixture=itp_01.xlsx call=BuildCallNoticeList(ITP, CallNotice) assert=COUNTA(ITP![Inspection Status]) = PRE_COUNTA(ITP![Inspection Status]) AND COUNTA(ITP![Activity]) = PRE_COUNTA(ITP![Activity])
TEST     fixture=itp_01.xlsx call=BuildCallNoticeList(CallNotice, CallNotice) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(CallNotice![Activity]) = PRE_COUNTA(CallNotice![Activity])
