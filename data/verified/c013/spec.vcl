VCL 0.6
TASK
# ASSUME: WPS register headers in row 1 include "WPS No", "PQR Ref" and "Review Date"
# ASSUME: 12 months before today is taken as 365 days before today, as VCL has no month arithmetic
# ASSUME: a PQR Ref is blank when it is empty or only spaces
# ASSUME: a Review Date that is blank or not a real date does not count as old
# ASSUME: only WPS No cells are cleared and highlighted; a row with a blank WPS No is skipped
GOAL     Highlight in yellow the WPS No of every WPS with a blank PQR Ref or a Review Date over 12 months old, clearing old highlights first, and return how many were highlighted.
PROC     HighlightWpsDue
KIND     Function
MODULE   modWps
ARGS     ws As Worksheet
RET      Long
VAR      cWps As Long
VAR      cPqr As Long
VAR      cReview As Long
VAR      due As Boolean
VAR      marked As Long
PRE      COL($ws, "WPS No") > 0; COL($ws, "PQR Ref") > 0; COL($ws, "Review Date") > 0
MODIFIES $ws fills
OP       LET(cWps, COL($ws, "WPS No"))
OP       LET(cPqr, COL($ws, "PQR Ref"))
OP       LET(cReview, COL($ws, "Review Date"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cWps) var=r
OP         FORMAT(CELL($ws, r, cWps), fill=none)
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cWps))
OP         LET(due, IS_EMPTY(CELL($ws, r, cPqr)))
OP         IF IS_DATE(CELL($ws, r, cReview))
OP           IF CELL($ws, r, cReview) < TODAY - 365
OP             LET(due, TRUE)
OP         IF due
OP           FORMAT(CELL($ws, r, cWps), fill=yellow)
OP           LET(marked, marked + 1)
OP       RETURN(marked)
PERF     screen_off calc_manual
ERR      standard
FIXTURE  wps_01.xlsx sheet=Wps cols=A:WPS No,B:Process,C:PQR Ref,D:Review Date rows=8 stale=2 data=A:text(6)+boundary("W-01","W-02","W-03","W-04","W-05","","W-06","W-07");B:values("SMAW"x4,"GTAW"x4);C:values("PQR-1"x2,"PQR-2"x2,"PQR-3"x1,""x3)+boundary("PQR-1","PQR-1","PQR-2","","","","PQR-2","PQR-3");D:reldates(-300,-10)+boundary(-366,-365,-364,-30,-400,-500,"",-10)
FIXTURE  wps_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=wps_01.xlsx call=HighlightWpsDue(Wps) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND FILL_COUNT(Wps![WPS No], #FFFF00) = 3 AND FILL_COUNT(Wps![WPS No], #FF0000) = 0
TEST     fixture=wps_01.xlsx call=HighlightWpsDue(Wps) assert=RESULT = 3 AND COUNTA(Wps![WPS No]) = PRE_COUNTA(Wps![WPS No]) AND COUNTA(Wps![Review Date]) = PRE_COUNTA(Wps![Review Date]) AND FILL_COUNT(Wps![WPS No]2, #FFFF00) = 1 AND FILL_COUNT(Wps![WPS No]3, #FFFF00) = 0 AND FILL_COUNT(Wps![WPS No]4, #FFFF00) = 0 AND FILL_COUNT(Wps![WPS No]5, #FFFF00) = 1 AND FILL_COUNT(Wps![WPS No]6, #FFFF00) = 1 AND FILL_COUNT(Wps![WPS No]7, #FFFF00) = 0 AND FILL_COUNT(Wps![WPS No]8, #FFFF00) = 0 AND FILL_COUNT(Wps![WPS No]9, #FFFF00) = 0
TEST     fixture=wps_01.xlsx call=HighlightWpsDue(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
