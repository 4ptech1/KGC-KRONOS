VCL 0.6
TASK
# ASSUME: WPS Register headers in row 1 include "WPS No", "Process", "Thickness Min", "Thickness Max" and "Remark"
# ASSUME: a thickness counts when it is a number or text that is a plain number; more than ten times means Max is greater than Min * 10, so exactly ten times is not marked
# ASSUME: rows where either thickness is blank, zero, negative or not a number keep their Remark as it is
# ASSUME: a row with both thicknesses valid and in range has its Remark cleared
# ASSUME: the macro returns the number of rows marked Check range
GOAL     Write Check range in Remark where Thickness Max is more than ten times Thickness Min, clear it on other valid rows and return how many were marked.
PROC     MarkWpsRanges
KIND     Function
MODULE   modWps
ARGS     ws As Worksheet
RET      Long
VAR      cMin As Long
VAR      cMax As Long
VAR      cRemark As Long
VAR      marked As Long
PRE      COL($ws, "WPS No") > 0; COL($ws, "Process") > 0; COL($ws, "Thickness Min") > 0; COL($ws, "Thickness Max") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Remark"]
OP       LET(cMin, COL($ws, "Thickness Min"))
OP       LET(cMax, COL($ws, "Thickness Max"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, COL($ws, "WPS No")) var=r
OP         SKIP_IF NOT IS_NUMBER(CELL($ws, r, cMin))
OP         SKIP_IF NOT IS_NUMBER(CELL($ws, r, cMax))
OP         SKIP_IF CELL($ws, r, cMin) <= 0
OP         SKIP_IF CELL($ws, r, cMax) <= 0
OP         IF CELL($ws, r, cMax) > CELL($ws, r, cMin) * 10
OP           SET_VALUE(CELL($ws, r, cRemark), "Check range")
OP           LET(marked, marked + 1)
OP         ELSE
OP           CLEAR(CELL($ws, r, cRemark), what=contents)
OP       RETURN(marked)
PERF     screen_off
ERR      standard
FIXTURE  wps_01.xlsx sheet=Wps cols=A:WPS No,B:Process,C:Thickness Min,D:Thickness Max,E:Remark rows=12 data=A:seq(1,1);B:values("GTAW"x6,"SMAW"x6);C:values("1"x4,"2"x2,""x1,"0"x1,"abc"x1,"3"x2,"4"x1)+boundary("1","1","2","2","1","1","","0","abc","3","3","4");D:int(5,50)+boundary(10,11,20,21,20,10,5,5,5,"",0,50);E:values("x"x12)
FIXTURE  wps_02.xlsx sheet=Wps cols=A:WPS No,B:Process,C:Thickness Min,D:Thickness Max,E:Remark rows=4 data=A:seq(1,1);B:values("GTAW"x4);C:int(1,2)+boundary(1,1,2,2);D:values("10"x1,"11"x1,"20"x1,"21"x1)+boundary("10","11","20","21");E:values("x"x4)
FIXTURE  wps_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=wps_01.xlsx call=MarkWpsRanges(Wps) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTIF(Wps![Remark], "Check range") = 4 AND COUNTIF(Wps![Remark], "x") = 5
TEST     fixture=wps_01.xlsx call=MarkWpsRanges(Wps) assert=VALUE(Wps![Remark]2) = "" AND VALUE(Wps![Remark]3) = "Check range" AND VALUE(Wps![Remark]4) = "" AND VALUE(Wps![Remark]5) = "Check range" AND VALUE(Wps![Remark]6) = "Check range" AND VALUE(Wps![Remark]7) = "" AND VALUE(Wps![Remark]13) = "Check range" AND COUNTA(Wps![WPS No]) = PRE_COUNTA(Wps![WPS No]) AND COUNTA(Wps![Thickness Max]) = PRE_COUNTA(Wps![Thickness Max])
TEST     fixture=wps_02.xlsx call=MarkWpsRanges(Wps) assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(Wps![Remark]2) = "" AND VALUE(Wps![Remark]3) = "Check range" AND VALUE(Wps![Remark]4) = "" AND VALUE(Wps![Remark]5) = "Check range"
TEST     fixture=wps_01.xlsx call=MarkWpsRanges(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
