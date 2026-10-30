VCL 0.6
TASK
# ASSUME: Welder Quals headers in row 1 include "Welder ID" and "Expiry Date"; the real sheet is "Welder Quals", the TESTs use WelderQuals
# ASSUME: red means the expiry date is before today; yellow means it is today or within the next 60 days (today + 60 included)
# ASSUME: a row with no date, or a date typed as text, gets no highlight
# ASSUME: the whole row means columns 1 to the last heading in row 1; old fills are cleared on that same span of every data row
GOAL     Highlight each welder qualification row red when it expired before today or yellow when it expires within 60 days, and return how many rows were highlighted.
PROC     HighlightWelderQuals
KIND     Function
MODULE   modWelderQuals
ARGS     ws As Worksheet
RET      Long
VAR      cId As Long
VAR      cDate As Long
VAR      lastC As Long
VAR      marked As Long
PRE      COL($ws, "Welder ID") > 0; COL($ws, "Expiry Date") > 0
MODIFIES $ws fills
OP       LET(cId, COL($ws, "Welder ID"))
OP       LET(cDate, COL($ws, "Expiry Date"))
OP       LET(lastC, LASTCOL($ws, 1))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cId) var=r
OP         FORMAT(ROW_RANGE($ws, r, 1, lastC), fill=none)
OP         IF IS_DATE(CELL($ws, r, cDate))
OP           IF CELL($ws, r, cDate) < TODAY
OP             FORMAT(ROW_RANGE($ws, r, 1, lastC), fill=red)
OP             LET(marked, marked + 1)
OP           ELSE
OP             IF CELL($ws, r, cDate) <= TODAY + 60
OP               FORMAT(ROW_RANGE($ws, r, 1, lastC), fill=yellow)
OP               LET(marked, marked + 1)
OP       RETURN(marked)
PERF     screen_off
ERR      standard
FIXTURE  welder_quals_01.xlsx sheet=WelderQuals cols=A:Welder ID,B:Name,C:Process,D:Expiry Date,E:Remark rows=10 stale=3 data=A:seq(1,1);B:text(8);C:values("GTAW"x5,"SMAW"x5);D:reldates(-90,150)+boundary(-1,0,60,61,"",-30,30,100,-5,45);E:text(6)+boundary("")
FIXTURE  welder_quals_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=welder_quals_01.xlsx call=HighlightWelderQuals(WelderQuals) runs=2 assert=RESULT = 7 AND ERROR_COUNT = 0 AND FILL_COUNT(WelderQuals![Welder ID], #FF0000) = 3 AND FILL_COUNT(WelderQuals![Welder ID], #FFFF00) = 4 AND FILL_COUNT(WelderQuals![Expiry Date], #FF0000) = 3 AND FILL_COUNT(WelderQuals![Expiry Date], #FFFF00) = 4 AND FILL_COUNT(WelderQuals![Remark], #FF0000) = 3 AND FILL_COUNT(WelderQuals![Remark], #FFFF00) = 4
TEST     fixture=welder_quals_01.xlsx call=HighlightWelderQuals(WelderQuals) assert=COUNTA(WelderQuals![Welder ID]) = PRE_COUNTA(WelderQuals![Welder ID]) AND COUNTA(WelderQuals![Expiry Date]) = PRE_COUNTA(WelderQuals![Expiry Date]) AND FILL_COUNT(WelderQuals![Name], #FF0000) = 3 AND FILL_COUNT(WelderQuals![Process], #FFFF00) = 4
TEST     fixture=welder_quals_01.xlsx call=HighlightWelderQuals(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
