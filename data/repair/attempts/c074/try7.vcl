VCL 0.6
TASK
# ASSUME: WPS Register headers in row 1 include "WPS No", "Thickness Min", "Thickness Max" and "Remark"
# ASSUME: a thickness counts as a number when IS_NUMBER accepts it, so numbers typed as text compare as numbers
# ASSUME: equal minimum and maximum is fine; a row is marked only when the minimum is bigger than the maximum
# ASSUME: every row up to the last WPS No is checked, and the Remark of an unmarked row is cleared
GOAL     Write Range? in Remark for every WPS whose thickness limits are missing, not numbers or reversed, clear the other remarks and return how many were marked.
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
OP         IF IS_NUMBER(CELL($ws, r, cMin))
OP           IF IS_NUMBER(CELL($ws, r, cMax))
OP             IF CELL($ws, r, cMin) > CELL($ws, r, cMax)
OP               SET_VALUE(CELL($ws, r, cRemark), "Range?")
OP               LET(marked, marked + 1)
OP             ELSE
OP               SET_VALUE(CELL($ws, r, cRemark), "")
OP           ELSE
OP             SET_VALUE(CELL($ws, r, cRemark), "Range?")
OP             LET(marked, marked + 1)
OP         ELSE
OP           SET_VALUE(CELL($ws, r, cRemark), "Range?")
OP           LET(marked, marked + 1)
OP       RETURN(marked)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  wps_01.xlsx sheet=Wps cols=A:WPS No,B:Process,C:Thickness Min,D:Thickness Max,E:Remark rows=8 data=A:seq(1,1);B:values("GTAW"x4,"SMAW"x4);C:values(6x2,12x1,""x1,"abc"x1,"6"x1,10x2)+boundary(6,12,6,"","abc","6",10,10);D:values(12x1,6x1,"12"x1,""x1,"n/a"x1,10x1,6x1,20x1)+boundary(12,6,6,12,"",12,"n/a",10);E:values("Range?"x3,"old"x2,""x3)+boundary("old","Range?","Range?","","","","Range?","old")
FIXTURE  wps_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=wps_01.xlsx call=MarkWpsRanges(Wps) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTIF(Wps![Remark], "Range?") = 5 AND COUNTIF(Wps![Remark], "old") = 0
TEST     fixture=wps_01.xlsx call=MarkWpsRanges(Wps) assert=VALUE(Wps![Remark]2) = "" AND VALUE(Wps![Remark]3) = "Range?" AND VALUE(Wps![Remark]4) = "" AND VALUE(Wps![Remark]5) = "Range?" AND VALUE(Wps![Remark]6) = "Range?" AND VALUE(Wps![Remark]7) = "Range?" AND VALUE(Wps![Remark]8) = "" AND VALUE(Wps![Remark]9) = "Range?"
TEST     fixture=wps_01.xlsx call=MarkWpsRanges(Wps) assert=COUNTA(Wps![WPS No]) = PRE_COUNTA(Wps![WPS No]) AND COUNTA(Wps![Thickness Min]) = PRE_COUNTA(Wps![Thickness Min]) AND COUNTA(Wps![Process]) = PRE_COUNTA(Wps![Process])
TEST     fixture=wps_01.xlsx call=MarkWpsRanges(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
