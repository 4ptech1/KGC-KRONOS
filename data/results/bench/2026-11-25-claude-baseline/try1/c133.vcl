VCL 0.7
TASK
# ASSUME: the WPS Register headings are in row 1 and the Remark heading already exists
# ASSUME: a thickness counts only when it is a number above zero (numbers typed as text count); other rows are left alone
# ASSUME: Check range is written when Thickness Max is more than ten times Thickness Min; exactly ten times is fine
GOAL     Mark each WPS whose maximum thickness is more than ten times its minimum, clear the Remark on the other valid rows, and count the rows marked.
PROC     MarkWpsRanges
KIND     Function
MODULE   modWps
ARGS     ws As Worksheet
RET      Long
VAR      cWps As Long
VAR      cMin As Long
VAR      cMax As Long
VAR      cRemark As Long
VAR      marked As Long
PRE      COL($ws, "WPS No") > 0; COL($ws, "Thickness Min") > 0; COL($ws, "Thickness Max") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Remark"]
OP       LET(cWps, COL($ws, "WPS No"))
OP       LET(cMin, COL($ws, "Thickness Min"))
OP       LET(cMax, COL($ws, "Thickness Max"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cWps) var=r
OP         SKIP_IF NOT IS_NUMBER(CELL($ws, r, cMin))
OP         SKIP_IF NOT IS_NUMBER(CELL($ws, r, cMax))
OP         SKIP_IF CELL($ws, r, cMin) <= 0
OP         SKIP_IF CELL($ws, r, cMax) <= 0
OP         IF CELL($ws, r, cMax) > 10 * CELL($ws, r, cMin)
OP           SET_VALUE(CELL($ws, r, cRemark), "Check range")
OP           LET(marked, marked + 1)
OP         ELSE
OP           CLEAR(CELL($ws, r, cRemark), what=contents)
OP       RETURN(marked)
ERR      standard
FIXTURE  wps_01.xlsx sheet=Wps cols=A:WPS No,B:Process,C:Thickness Min,D:Thickness Max,E:Remark rows=8 data=A:seq(1,1);B:text(4);C:values("1"x2,"2"x1,""x1,"0"x1,"n/a"x1,"5"x1,"3"x1)+boundary("1","1","2","","0","n/a","5","3");D:values("10"x1,"11"x1,"30"x1,"50"x3,""x1,"20"x1)+boundary("10","11","30","50","50","50","","20");E:values("old"x8)
FIXTURE  wps_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=wps_01.xlsx call=MarkWpsRanges(Wps) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND COUNTIF(Wps![Remark], "Check range") = 2 AND COUNTIF(Wps![Remark], "old") = 4 AND VALUE(Wps![Remark]2) = "" AND VALUE(Wps![Remark]3) = "Check range" AND VALUE(Wps![Remark]9) = ""
TEST     fixture=wps_01.xlsx call=MarkWpsRanges(Wps) assert=COUNTA(Wps![WPS No]) = PRE_COUNTA(Wps![WPS No]) AND COUNTA(Wps![Thickness Min]) = PRE_COUNTA(Wps![Thickness Min]) AND COUNTA(Wps![Thickness Max]) = PRE_COUNTA(Wps![Thickness Max])
TEST     fixture=wps_01.xlsx call=MarkWpsRanges(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
