VCL 0.6
TASK
# ASSUME: Welder Qualification headers (row 1) include "Welder ID" and "Expiry Date"; one row per welder
# ASSUME: Weld Log headers include "Weld Seam No" and "Welder ID"; the Weld Log row is the one highlighted
# ASSUME: a Weld Log row is red when its Welder ID is blank, not in the qualification list, has no valid expiry date, or the expiry date is before today
# ASSUME: only the expiry date is checked, not whether the welder is qualified for that specific WPS
# ASSUME: the macro owns the fill colour of Weld Log rows: it clears the fill first, so fixed rows lose their red
GOAL     Highlight in red every weld log row whose welder is missing from the qualification list or has no valid or an expired expiry date.
PROC     FlagExpiredWelderQualifications
KIND     Function
MODULE   modWelding
ARGS     weldLog As Worksheet, qual As Worksheet
RET      Long
VAR      expiry As Dictionary
VAR      cSeam As Long
VAR      cWelder As Long
VAR      cQualId As Long
VAR      cExpiry As Long
VAR      lastCol As Long
VAR      flagged As Long
VAR      welderId As String
VAR      expDate As Variant
VAR      bad As Boolean
PRE      COL($weldLog, "Weld Seam No") > 0; COL($weldLog, "Welder ID") > 0; COL($qual, "Welder ID") > 0; COL($qual, "Expiry Date") > 0
MODIFIES $weldLog fills
OP       LET(cSeam, COL($weldLog, "Weld Seam No"))
OP       LET(cWelder, COL($weldLog, "Welder ID"))
OP       LET(cQualId, COL($qual, "Welder ID"))
OP       LET(cExpiry, COL($qual, "Expiry Date"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($qual, cQualId) var=r
OP         IF NOT IS_EMPTY(CELL($qual, r, cQualId)) THEN DICT_ADD(expiry, TRIM(CELL($qual, r, cQualId)), CELL($qual, r, cExpiry), mode=overwrite)
OP       LET(lastCol, LASTCOL($weldLog, 1))
OP       LET(flagged, 0)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($weldLog, cSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($weldLog, r, cSeam))
OP         FORMAT(ROW_RANGE($weldLog, r, 1, lastCol), fill=none)
OP         LET(bad, FALSE)
OP         LET(expDate, "")
OP         LET(welderId, TRIM(CELL($weldLog, r, cWelder)))
OP         IF IS_EMPTY(welderId) THEN LET(bad, TRUE)
OP         IF EXISTS(expiry, welderId) THEN DICT_GET(expiry, welderId, into=expDate)
OP         IF NOT IS_DATE(expDate) THEN LET(bad, TRUE)
OP         IF IS_DATE(expDate)
OP           IF expDate < TODAY THEN LET(bad, TRUE)
OP         IF bad
OP           FORMAT(ROW_RANGE($weldLog, r, 1, lastCol), fill=#FF0000)
OP           LET(flagged, flagged + 1)
OP       RETURN(flagged)
PERF     screen_off
ERR      standard
FIXTURE  welder_qual_01.xlsx sheet=WelderQual cols=A:Welder ID,B:Name,C:Expiry Date rows=7 data=A:values("W-01"x2,"W-02"x1,"W-03"x1,"W-04"x1,"W-05"x1,"W-06"x1)+boundary("W-02","W-01","W-03","W-01","W-04");B:text(10);C:reldates(-400,800)+boundary(0,-1,1,30,"")
FIXTURE  welder_qual_01.xlsx sheet=WeldLog cols=A:Weld Seam No,B:Welder ID,C:Fit-up,D:Visual rows=40 stale=5 data=A:seq(1,1)+blank(2);B:values("W-01"x8,"W-02"x8,"W-03"x8,"W-04"x8,"W-07"x6,""x2)+boundary("W-07");C:values("Accepted"x40);D:values("Accepted"x39,""x1)+boundary("")
TEST     fixture=welder_qual_01.xlsx call=FlagExpiredWelderQualifications(WeldLog, WelderQual) runs=2 assert=RESULT = 15 AND FILL_COUNT(WeldLog![Weld Seam No], #FF0000) = 15 AND FILL_COUNT(WeldLog![Visual], #FF0000) = 15 AND FILL_COUNT(WeldLog![Welder ID]2, #FF0000) = 1 AND ERROR_COUNT = 0
TEST     fixture=welder_qual_01.xlsx call=FlagExpiredWelderQualifications(WeldLog, WelderQual) assert=COUNTA(WeldLog![Welder ID]) = PRE_COUNTA(WeldLog![Welder ID]) AND COUNTA(WelderQual![Expiry Date]) = PRE_COUNTA(WelderQual![Expiry Date])
TEST     fixture=welder_qual_01.xlsx call=FlagExpiredWelderQualifications(WelderQual, WelderQual) assert=RESULT = -1 AND ERROR_COUNT = 1 AND FILL_COUNT(WelderQual![Welder ID], #FF0000) = 0
