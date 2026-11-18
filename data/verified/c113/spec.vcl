VCL 0.6
TASK
# ASSUME: Material Certs headers in row 1 include "Heat No", "MTC No", "Grade" and "Received Date"
# ASSUME: more than 14 days ago means Received Date is before TODAY - 14; exactly 14 days ago is not highlighted
# ASSUME: a row whose Received Date is blank or not a real date is not highlighted
# ASSUME: the highlight covers the row from column A to the last used column, and every row up to the last Heat No is checked
# ASSUME: all fills on the sheet are cleared first, since the highlight is the only fill the sheet should carry
GOAL     Clear old row highlights in Material Certs, then highlight yellow every row with no MTC No received more than 14 days ago, and return how many were highlighted.
PROC     HighlightMissingCerts
KIND     Function
MODULE   modCerts
ARGS     ws As Worksheet
RET      Long
VAR      cHeat As Long
VAR      cMtc As Long
VAR      cDate As Long
VAR      lastCol As Long
VAR      marked As Long
PRE      COL($ws, "Heat No") > 0; COL($ws, "MTC No") > 0; COL($ws, "Received Date") > 0
MODIFIES $ws fills
OP       LET(cHeat, COL($ws, "Heat No"))
OP       LET(cMtc, COL($ws, "MTC No"))
OP       LET(cDate, COL($ws, "Received Date"))
OP       LET(lastCol, LASTCOL($ws, 1))
OP       FORMAT(USED_RANGE($ws), fill=none)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cHeat) var=r
OP         IF IS_EMPTY(CELL($ws, r, cMtc))
OP           IF IS_DATE(CELL($ws, r, cDate))
OP             IF CELL($ws, r, cDate) < TODAY - 14
OP               FORMAT(ROW_RANGE($ws, r, 1, lastCol), fill=yellow)
OP               LET(marked, marked + 1)
OP       RETURN(marked)
PERF     screen_off
ERR      standard
FIXTURE  certs_01.xlsx sheet=Certs cols=A:Heat No,B:MTC No,C:Grade,D:Received Date rows=9 stale=3 data=A:seq(1,1);B:values(""x6,"M-1"x3)+boundary("","","","M-1","","","M-1","","M-1");C:values("SA-516"x5,"SA-106"x4);D:reldates(-40,-1)+boundary(-15,-14,-13,-30,"",-40,-15,-15,-2)
FIXTURE  certs_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=certs_01.xlsx call=HighlightMissingCerts(Certs) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND FILL_COUNT(Certs![Heat No], #FFFF00) = 3 AND FILL_COUNT(Certs![Received Date], #FFFF00) = 3 AND FILL_COUNT(Certs![MTC No], #FFFF00) = 3 AND FILL_COUNT(Certs![Grade], #FFFF00) = 3 AND FILL_COUNT(Certs![Heat No], #FF0000) = 0
TEST     fixture=certs_01.xlsx call=HighlightMissingCerts(Certs) assert=FILL_COUNT(Certs![Heat No]2, #FFFF00) = 1 AND FILL_COUNT(Certs![Heat No]3, #FFFF00) = 0 AND FILL_COUNT(Certs![Heat No]4, #FFFF00) = 0 AND FILL_COUNT(Certs![Heat No]5, #FFFF00) = 0 AND FILL_COUNT(Certs![Heat No]6, #FFFF00) = 0 AND FILL_COUNT(Certs![Heat No]7, #FFFF00) = 1 AND FILL_COUNT(Certs![Heat No]8, #FFFF00) = 0 AND FILL_COUNT(Certs![Heat No]9, #FFFF00) = 1 AND FILL_COUNT(Certs![Heat No]10, #FFFF00) = 0
TEST     fixture=certs_01.xlsx call=HighlightMissingCerts(Certs) assert=COUNTA(Certs![Heat No]) = PRE_COUNTA(Certs![Heat No]) AND COUNTA(Certs![MTC No]) = PRE_COUNTA(Certs![MTC No]) AND COUNTA(Certs![Received Date]) = PRE_COUNTA(Certs![Received Date])
TEST     fixture=certs_01.xlsx call=HighlightMissingCerts(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
