VCL 0.7
TASK
# ASSUME: the register's headings are in row 1 and include "Enquiry" and "Result"
# ASSUME: an enquiry is a row whose Enquiry cell holds a value (the top-left cell of its merged block); Result is read in that same row
# ASSUME: a blank Result means empty or only spaces; the highlight is yellow on the Enquiry cell only
# ASSUME: old highlights are the fills on the whole Enquiry column below the heading, all removed before the new ones are made
# ASSUME: the second practice fixture is a plain sheet because the template has no old fills to clear
GOAL     Clear the fills of the Enquiry column and highlight yellow the Enquiry cell of every enquiry whose Result is still blank, returning how many.
PROC     HighlightOpenEnquiries
KIND     Function
MODULE   modRegister
ARGS     ws As Worksheet
RET      Long
VAR      cEnq As Long
VAR      cRes As Long
VAR      marked As Long
PRE      COL($ws, "Enquiry") > 0; COL($ws, "Result") > 0
MODIFIES $ws fills
OP       LET(cEnq, COL($ws, "Enquiry"))
OP       LET(cRes, COL($ws, "Result"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cEnq) var=r
OP         FORMAT(CELL($ws, r, cEnq), fill=none)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cEnq) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cEnq))
OP         IF IS_EMPTY(CELL($ws, r, cRes))
OP           FORMAT(CELL($ws, r, cEnq), fill=yellow)
OP           LET(marked, marked + 1)
OP       RETURN(marked)
PERF     screen_off
ERR      standard
FIXTURE  register_01.xlsx template=enquiry_register
FIXTURE  plain_01.xlsx sheet=Sheet1 cols=A:Enquiry,B:Result rows=6 stale=2 data=A:seq(1,1)+boundary("");B:values("Won"x2,"Lost"x1,""x3)+boundary("","","")
FIXTURE  plain_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=register_01.xlsx call=HighlightOpenEnquiries(Sheet1) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND FILL_COUNT(Sheet1![Enquiry], #FFFF00) = 2 AND FILL_COUNT(Sheet1![Enquiry], #FFFF00) = RESULT
TEST     fixture=register_01.xlsx call=HighlightOpenEnquiries(Sheet1) assert=COUNTA(Sheet1![Enquiry]) = PRE_COUNTA(Sheet1![Enquiry]) AND COUNTA(Sheet1![Result]) = PRE_COUNTA(Sheet1![Result]) AND FILL_COUNT(Sheet1![Result], #FFFF00) = 0
TEST     fixture=plain_01.xlsx call=HighlightOpenEnquiries(Sheet1) assert=RESULT = 2 AND ERROR_COUNT = 0 AND FILL_COUNT(Sheet1![Enquiry], #FF0000) = 0 AND FILL_COUNT(Sheet1![Enquiry], #FFFF00) = 2
TEST     fixture=plain_01.xlsx call=HighlightOpenEnquiries(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
