VCL 0.6
TASK
# ASSUME: Weld Log headers in row 1 include "Weld Seam No" and "Welder ID"; Welder Quals headers include "Welder ID" and "Name"
# ASSUME: the real sheets are "Weld Log", "Welder Quals" and "Unknown Welders"; the TESTs use WeldLog, WelderQuals and UnknownWelders
# ASSUME: the Unknown Welders sheet already exists; it is rebuilt on every run with the header Welder ID
# ASSUME: the Weld Seam No and Name headers are only checked in PRE, so the wrong sheet is refused
# ASSUME: Welder IDs are cleaned before matching (spaces at the ends and inside runs of spaces) and compared ignoring case
# ASSUME: blank Welder IDs are skipped; unknown IDs are listed in the order they first appear in the weld log
GOAL     List on the Unknown Welders sheet each Welder ID used in the weld log that is missing from Welder Quals, one row per ID.
PROC     ListUnknownWelders
KIND     Function
MODULE   modWelders
ARGS     weldLog As Worksheet, quals As Worksheet, out As Worksheet
RET      Long
VAR      cLog As Long
VAR      cQual As Long
VAR      id As String
VAR      known As Dictionary
VAR      unknown As Dictionary
VAR      outRow As Long
PRE      COL($weldLog, "Weld Seam No") > 0; COL($weldLog, "Welder ID") > 0; COL($quals, "Welder ID") > 0; COL($quals, "Name") > 0
MODIFIES $out
OP       LET(cLog, COL($weldLog, "Welder ID"))
OP       LET(cQual, COL($quals, "Welder ID"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($quals, cQual) var=r
OP         SKIP_IF IS_EMPTY(CELL($quals, r, cQual))
OP         DICT_ADD(known, CLEAN_TEXT(CELL($quals, r, cQual)), 1, mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($weldLog, cLog) var=r
OP         SKIP_IF IS_EMPTY(CELL($weldLog, r, cLog))
OP         LET(id, CLEAN_TEXT(CELL($weldLog, r, cLog)))
OP         IF NOT EXISTS(known, id) THEN DICT_ADD(unknown, id, 1, mode=keep)
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Welder ID")
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(unknown)
OP         WRITE_ROW($out, outRow, k)
OP         LET(outRow, outRow + 1)
OP       RETURN(COUNT(unknown))
PERF     screen_off
ERR      standard
FIXTURE  welders_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Welder ID rows=12 data=A:seq(1,1);B:values("W-01"x1," W-01 "x1,"W-99"x1,"w-02"x1," W-99"x1,"W-77"x1,""x1,"W-03"x1,"W  98"x1,"W 98"x1,"W-77"x1,"W-04"x1)+boundary("W-01"," W-01 ","W-99","w-02"," W-99","W-77","","W-03","W  98","W 98","W-77","W-04")
FIXTURE  welders_01.xlsx sheet=WelderQuals cols=A:Welder ID,B:Name rows=5 data=A:values("W-01"x1," W-02"x1,"W-03 "x1,""x1,"W-05"x1)+boundary("W-01"," W-02","W-03 ","","W-05");B:text(8)
FIXTURE  welders_01.xlsx sheet=UnknownWelders cols=A:Welder ID rows=0 stale=2
TEST     fixture=welders_01.xlsx call=ListUnknownWelders(WeldLog, WelderQuals, UnknownWelders) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND COUNTA(UnknownWelders![Welder ID]) = 5 AND COUNTIF(UnknownWelders![Welder ID], "stale") = 0 AND VALUE(UnknownWelders![Welder ID]2) = "W-99" AND VALUE(UnknownWelders![Welder ID]3) = "W-77" AND VALUE(UnknownWelders![Welder ID]4) = "W 98" AND VALUE(UnknownWelders![Welder ID]5) = "W-04"
TEST     fixture=welders_01.xlsx call=ListUnknownWelders(WeldLog, WelderQuals, UnknownWelders) assert=COUNTIF(UnknownWelders![Welder ID], "W-01") = 0 AND COUNTIF(UnknownWelders![Welder ID], "W-02") = 0 AND COUNTIF(UnknownWelders![Welder ID], "W-03") = 0 AND COUNTIF(UnknownWelders![Welder ID], "W-99") = 1 AND COUNTA(WeldLog![Welder ID]) = PRE_COUNTA(WeldLog![Welder ID]) AND COUNTA(WelderQuals![Welder ID]) = PRE_COUNTA(WelderQuals![Welder ID]) AND COUNTA(WelderQuals![Name]) = PRE_COUNTA(WelderQuals![Name])
TEST     fixture=welders_01.xlsx call=ListUnknownWelders(WelderQuals, WeldLog, UnknownWelders) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(UnknownWelders![Welder ID]) = PRE_COUNTA(UnknownWelders![Welder ID])
