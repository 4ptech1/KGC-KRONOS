VCL 0.6
TASK
# ASSUME: WPS Register headers in row 1 include "WPS No", "Thickness Min", "Thickness Max" and "Remark"
# ASSUME: Range? is written when either thickness is blank or not a number, or when the minimum is bigger than the maximum; equal values are fine
# ASSUME: the Remark is cleared on every other WPS, whatever it held before; the last row is found from WPS No
GOAL     Write Range? in Remark for each WPS with a missing, non-numeric or inverted thickness range, clear the Remark on the others and return how many were marked.
PROC     MarkBadThicknessRanges
KIND     Function
MODULE   modWps
ARGS     ws As Worksheet
RET      Long
VAR      cWps As Long
VAR      cMin As Long
VAR      cMax As Long
VAR      cRemark As Long
VAR      bad As Long
VAR      marked As Long
PRE      COL($ws, "WPS No") > 0; COL($ws, "Thickness Min") > 0; COL($ws, "Thickness Max") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Remark"]
OP       LET(cWps, COL($ws, "WPS No"))
OP       LET(cMin, COL($ws, "Thickness Min"))
OP       LET(cMax, COL($ws, "Thickness Max"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cWps) var=r
OP         LET(bad, 1)
OP         IF IS_NUMBER(CELL($ws, r, cMin))
OP           IF IS_NUMBER(CELL($ws, r, cMax))
OP             IF CELL($ws, r, cMax) - CELL($ws, r, cMin) >= 0 THEN LET(bad, 0)
OP         IF bad = 1
OP           SET_VALUE(CELL($ws, r, cRemark), "Range?")
OP           LET(marked, marked + 1)
OP         ELSE
OP           CLEAR(CELL($ws, r, cRemark), what=contents)
OP       RETURN(marked)
ERR      standard
FIXTURE  wps_01.xlsx sheet=Wps cols=A:WPS No,B:Process,C:Thickness Min,D:Thickness Max,E:Remark rows=10 data=A:seq(1,1);B:values("GTAW"x5,"SMAW"x5);C:values("6"x1,"12"x1,"13"x1,""x1,"6"x1,"n/a"x1,"6"x1,"9"x1,"10"x1,"3"x1)+boundary("6","12","13","","6","n/a","6","9","10","3");D:values("12"x4,""x1,"tbc"x1,"10"x1,"9"x1,"20"x1)+boundary("12","12","12","12","","12","tbc","10","9","20");E:values("old"x6,"Range?"x4)+boundary("Range?","old","old","Range?","old","old","Range?","old","old","Range?")
FIXTURE  wps_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=wps_01.xlsx call=MarkBadThicknessRanges(Wps) runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND COUNTIF(Wps![Remark], "Range?") = 6 AND COUNTIF(Wps![Remark], "old") = 0
TEST     fixture=wps_01.xlsx call=MarkBadThicknessRanges(Wps) assert=VALUE(Wps![Remark]2) = "" AND VALUE(Wps![Remark]3) = "" AND VALUE(Wps![Remark]4) = "Range?" AND VALUE(Wps![Remark]9) = "" AND VALUE(Wps![Remark]10) = "Range?" AND VALUE(Wps![Remark]11) = ""
TEST     fixture=wps_01.xlsx call=MarkBadThicknessRanges(Wps) assert=COUNTA(Wps![WPS No]) = PRE_COUNTA(Wps![WPS No]) AND SUM(Wps![Thickness Min]) = PRE_SUM(Wps![Thickness Min]) AND COUNTA(Wps![Thickness Max]) = PRE_COUNTA(Wps![Thickness Max])
TEST     fixture=wps_01.xlsx call=MarkBadThicknessRanges(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
