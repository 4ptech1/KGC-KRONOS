VCL 0.6
TASK
# ASSUME: WPS Register headers in row 1 include "WPS No", "Thickness Min", "Thickness Max" and "Remark"
# ASSUME: a row with a blank WPS No is not a WPS: its Remark is left as it is and it is not counted
# ASSUME: a minimum equal to the maximum is a good range; a number typed as text counts as a number
# ASSUME: clearing a Remark removes whatever text it held, not only "Range?"
GOAL     Write Range? in the Remark of every WPS whose thickness limits are blank, not numbers or the wrong way round, clear the other Remarks and return how many were marked.
PROC     MarkBadThicknessRanges
KIND     Function
MODULE   modWps
ARGS     ws As Worksheet
RET      Long
VAR      cNo As Long
VAR      cMin As Long
VAR      cMax As Long
VAR      cRemark As Long
VAR      bad As Boolean
VAR      marked As Long
PRE      COL($ws, "WPS No") > 0; COL($ws, "Thickness Min") > 0; COL($ws, "Thickness Max") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Remark"]
OP       LET(cNo, COL($ws, "WPS No"))
OP       LET(cMin, COL($ws, "Thickness Min"))
OP       LET(cMax, COL($ws, "Thickness Max"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cNo) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cNo))
OP         LET(bad, TRUE)
OP         IF IS_NUMBER(CELL($ws, r, cMin))
OP           IF IS_NUMBER(CELL($ws, r, cMax))
OP             IF CELL($ws, r, cMin) <= CELL($ws, r, cMax) THEN LET(bad, FALSE)
OP         IF bad
OP           SET_VALUE(CELL($ws, r, cRemark), "Range?")
OP           LET(marked, marked + 1)
OP         ELSE
OP           CLEAR(CELL($ws, r, cRemark), what=contents)
OP       RETURN(marked)
ERR      standard
FIXTURE  wps_ranges_01.xlsx sheet=Register cols=A:WPS No,B:Process,C:Thickness Min,D:Thickness Max,E:Remark rows=11 data=A:values("W-01"x1,"W-02"x1,"W-03"x1,"W-04"x1,"W-05"x1,"W-06"x1,"W-07"x1,""x1,"W-09"x1,"W-10"x1,"W-11"x1)+boundary("W-01","W-02","W-03","W-04","W-05","W-06","W-07","","W-09","W-10","W-11");B:values("GTAW"x4,"SMAW"x4,"SAW"x3);C:int(1,20)+boundary(3,10,11,"",0,"abc",3,20,5,"4","");D:int(1,20)+boundary(10,10,10,10,"",10,"n/a",5,6,12,"");E:values("Range?"x3,"old"x3,"keep"x1,""x4)+boundary("Range?","","","old","Range?","","old","keep","old","Range?","")
FIXTURE  wps_ranges_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=wps_ranges_01.xlsx call=MarkBadThicknessRanges(Register) runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND COUNTIF(Register![Remark], "Range?") = 6 AND COUNTA(Register![Remark]) = 8 AND VALUE(Register![Remark]4) = "Range?" AND VALUE(Register![Remark]5) = "Range?" AND VALUE(Register![Remark]6) = "Range?" AND VALUE(Register![Remark]7) = "Range?" AND VALUE(Register![Remark]8) = "Range?" AND VALUE(Register![Remark]12) = "Range?" AND VALUE(Register![Remark]9) = "keep"
TEST     fixture=wps_ranges_01.xlsx call=MarkBadThicknessRanges(Register) assert=RESULT = 6 AND ERROR_COUNT = 0 AND COUNTA(Register![WPS No]) = PRE_COUNTA(Register![WPS No]) AND COUNTA(Register![Thickness Min]) = PRE_COUNTA(Register![Thickness Min]) AND COUNTA(Register![Thickness Max]) = PRE_COUNTA(Register![Thickness Max]) AND SUM(Register![Thickness Max]) = PRE_SUM(Register![Thickness Max]) AND COUNTIF(Register![Remark], "old") = 0
TEST     fixture=wps_ranges_01.xlsx call=MarkBadThicknessRanges(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
