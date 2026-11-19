VCL 0.6
TASK
# ASSUME: NDE Report headers in row 1 include "Weld Seam No", "Method", "Result" and "Report Date"
# ASSUME: Result is matched ignoring case and extra spaces; Report Date holds real dates
# ASSUME: more than 3 days ago means the Report Date is before today minus 3 days
# ASSUME: a Reject row is red even if its date is old; a Pending row with no real date stays uncoloured
# ASSUME: all old fills on the sheet are cleared first, header row included, then the colours are applied across columns 1 to the last used column
# ASSUME: the macro returns the number of red rows
GOAL     Clear old row colours in the NDE Report, colour rejected rows red and pending rows older than 3 days yellow, and return the number of red rows.
PROC     ColourNdeRows
KIND     Function
MODULE   modNde
ARGS     ws As Worksheet
RET      Long
VAR      cResult As Long
VAR      cDate As Long
VAR      lastCol As Long
VAR      red As Long
PRE      COL($ws, "Weld Seam No") > 0; COL($ws, "Method") > 0; COL($ws, "Result") > 0; COL($ws, "Report Date") > 0
MODIFIES $ws fills
OP       LET(cResult, COL($ws, "Result"))
OP       LET(cDate, COL($ws, "Report Date"))
OP       LET(lastCol, LASTCOL($ws, 1))
OP       FORMAT(USED_RANGE($ws), fill=none)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, COL($ws, "Weld Seam No")) var=r
OP         IF SAME(CELL($ws, r, cResult), "Reject")
OP           FORMAT(ROW_RANGE($ws, r, 1, lastCol), fill=red)
OP           LET(red, red + 1)
OP         ELSE
OP           IF SAME(CELL($ws, r, cResult), "Pending")
OP             IF IS_DATE(CELL($ws, r, cDate))
OP               IF CELL($ws, r, cDate) < TODAY - 3
OP                 FORMAT(ROW_RANGE($ws, r, 1, lastCol), fill=yellow)
OP       RETURN(red)
PERF     screen_off
ERR      standard
FIXTURE  nde_01.xlsx sheet=Nde cols=A:Weld Seam No,B:Method,C:Result,D:Report Date rows=12 stale=3 data=A:seq(1,1);B:values("RT"x6,"UT"x6);C:values("Reject"x3,"Pending"x5,"Accepted"x2,"reject "x1,"PENDING"x1)+boundary("Pending","Pending","Pending","Reject","reject ","Pending","PENDING","Accepted","Reject","Pending");D:reldates(-20,0)+boundary(-4,-3,-2,-10,-10,"",-5,-10,0,-7)
TEST     fixture=nde_01.xlsx call=ColourNdeRows(Nde) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND FILL_COUNT(Nde![Weld Seam No], #FF0000) = 3 AND FILL_COUNT(Nde![Method], #FF0000) = 3
TEST     fixture=nde_01.xlsx call=ColourNdeRows(Nde) assert=FILL_COUNT(Nde![Weld Seam No], #FFFF00) = 3 AND FILL_COUNT(Nde![Report Date], #FFFF00) = 3 AND FILL_COUNT(Nde![Result], #FFFF00) = 3 AND COUNTA(Nde![Result]) = PRE_COUNTA(Nde![Result]) AND COUNTA(Nde![Report Date]) = PRE_COUNTA(Nde![Report Date])
TEST     fixture=nde_01.xlsx call=ColourNdeRows(Nde) assert=FILL_COUNT(Nde![Weld Seam No]2, #FFFF00) = 0 AND FILL_COUNT(Nde![Weld Seam No]3, #FFFF00) = 0 AND FILL_COUNT(Nde![Weld Seam No]4, #FFFF00) = 1 AND FILL_COUNT(Nde![Weld Seam No]5, #FF0000) = 1 AND FILL_COUNT(Nde![Weld Seam No]6, #FF0000) = 1
