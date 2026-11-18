VCL 0.6
TASK
# ASSUME: the NCR sheet has the headers NCR No in column A and Status in column B; PRE checks this because a formula's text can't use COL()
# ASSUME: the NCR sheet is really named NCR; the new sheet is NCR Stats, added after the last sheet
# ASSUME: the cells are written only in the run that adds the sheet; when NCR Stats is already there nothing is changed
# ASSUME: the formula counts cells of column B of NCR that read Open (COUNTIF, so case is ignored); the heading Status is not counted
GOAL     Add an NCR Stats sheet if it is missing, with the label Open NCRs in A1 and a formula in B1 that counts the Open cells in the Status column of the NCR sheet.
PROC     AddNcrStatsSheet
KIND     Function
MODULE   modNcrStats
ARGS     book As Workbook, ncr As Worksheet
RET      Long
VAR      stats As Worksheet
PRE      COL($ncr, "NCR No") = 1; COL($ncr, "Status") = 2
MODIFIES $book sheets
OP       EXIT_IF SHEET_EXISTS($book, "NCR Stats")
OP       ADD_SHEET($book, name="NCR Stats", into=stats)
OP       SET_VALUE(CELL(stats, 1, 1), "Open NCRs")
OP       SET_FORMULA(CELL(stats, 1, 2), "=COUNTIF(NCR!B:B,""Open"")", style=a1)
OP       RETURN(1)
ERR      standard
FIXTURE  ncr_01.xlsx sheet=NCR cols=A:NCR No,B:Status rows=6 data=A:seq(1,1);B:values("Open"x3,"Closed"x2,"open "x1)
FIXTURE  ncr_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=ncr_01.xlsx call=AddNcrStatsSheet(ncr_01.xlsx, NCR) assert=RESULT = 1 AND ERROR_COUNT = 0 AND SHEET_EXISTS("NCR Stats") AND VALUE('NCR Stats'!A1) = "Open NCRs" AND FORMULA('NCR Stats'!B1) = "=COUNTIF(NCR!B:B,""Open"")" AND COUNTA('NCR Stats'!A1:B1) = 2
TEST     fixture=ncr_01.xlsx call=AddNcrStatsSheet(ncr_01.xlsx, NCR) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND SHEET_EXISTS("NCR Stats") AND COUNTA(NCR![NCR No]) = PRE_COUNTA(NCR![NCR No]) AND COUNTIF(NCR![Status], "Open") = PRE_COUNTIF(NCR![Status], "Open")
TEST     fixture=ncr_01.xlsx call=AddNcrStatsSheet(ncr_01.xlsx, Notes) assert=RESULT = -1 AND ERROR_COUNT = 1 AND SHEET_EXISTS("NCR Stats") = 0
