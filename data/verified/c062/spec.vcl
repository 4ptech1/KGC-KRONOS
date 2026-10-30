VCL 0.6
TASK
# ASSUME: the weld log sheet has headers in row 1 that include "Weld No" and "Process"
# ASSUME: a Process cell counts as blank when it is empty or holds only spaces; any other text is left alone
# ASSUME: the rows to check run down to the last row of the Weld No column
GOAL     Fill every blank Process cell in the weld log with TBC, leave cells that already have a process alone, and return how many were filled.
PROC     FillBlankProcess
KIND     Function
MODULE   modWeldLog
ARGS     ws As Worksheet
RET      Long
VAR      cWeld As Long
VAR      cProc As Long
VAR      filled As Long
PRE      COL($ws, "Weld No") > 0; COL($ws, "Process") > 0
MODIFIES $ws["Process"]
OP       LET(cWeld, COL($ws, "Weld No"))
OP       LET(cProc, COL($ws, "Process"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cWeld) var=r
OP         IF IS_EMPTY(CELL($ws, r, cProc))
OP           SET_VALUE(CELL($ws, r, cProc), "TBC")
OP           LET(filled, filled + 1)
OP       RETURN(filled)
ERR      standard
FIXTURE  weld_01.xlsx sheet=WeldLog cols=A:Weld No,B:Process,C:Welder rows=12 data=A:seq(1,1);B:values("SMAW"x3,"GTAW"x3,"FCAW"x1,"TBC"x1,""x3," "x1)+boundary("","SMAW"," ","GTAW","TBC","");C:text(6)
FIXTURE  weld_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=weld_01.xlsx call=FillBlankProcess(WeldLog) runs=2 assert=COUNTIF(WeldLog![Process], "TBC") = 5 AND COUNTA(WeldLog![Process]) = PRE_COUNTA(WeldLog![Process]) + 3 AND ERROR_COUNT = 0
TEST     fixture=weld_01.xlsx call=FillBlankProcess(WeldLog) assert=RESULT = 4 AND VALUE(WeldLog![Process]2) = "TBC" AND VALUE(WeldLog![Process]7) = "TBC" AND VALUE(WeldLog![Process]4) = "TBC" AND VALUE(WeldLog![Process]3) = "SMAW" AND VALUE(WeldLog![Process]5) = "GTAW" AND COUNTIF(WeldLog![Process], "SMAW") = PRE_COUNTIF(WeldLog![Process], "SMAW") AND COUNTIF(WeldLog![Process], "GTAW") = PRE_COUNTIF(WeldLog![Process], "GTAW") AND COUNTIF(WeldLog![Process], "FCAW") = PRE_COUNTIF(WeldLog![Process], "FCAW") AND COUNTA(WeldLog![Weld No]) = PRE_COUNTA(WeldLog![Weld No])
TEST     fixture=weld_01.xlsx call=FillBlankProcess(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
