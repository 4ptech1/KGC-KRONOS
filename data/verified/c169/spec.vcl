VCL 0.7
TASK
# ASSUME: the real sheet is "Enq Summary"; the TESTs use the short form EnqSummary with the template's own headings Sr.Nos., Enquiry, Date and Item
# ASSUME: an item row is a row whose Item cell is not blank; rows with a blank Item are left alone
# ASSUME: today's date is written as a real date and shown as dd.mm.yyyy through its number format
# ASSUME: the enquiry number is written as the text E-9100, replacing what the cell held
GOAL     Write enquiry number E-9100 and today's date as dd.mm.yyyy in the Enquiry and Date columns of every item row on the Enq Summary sheet.
PROC     StampEnquiry
KIND     Function
MODULE   modEnquiry
ARGS     ws As Worksheet
RET      Long
VAR      cEnq As Long
VAR      cDate As Long
VAR      cItem As Long
VAR      n As Long
PRE      COL($ws, "Enquiry") > 0; COL($ws, "Date") > 0; COL($ws, "Item") > 0
MODIFIES $ws
OP       LET(cEnq, COL($ws, "Enquiry"))
OP       LET(cDate, COL($ws, "Date"))
OP       LET(cItem, COL($ws, "Item"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cItem) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cItem))
OP         SET_VALUE(CELL($ws, r, cEnq), "E-9100")
OP         SET_VALUE(CELL($ws, r, cDate), TODAY)
OP         FORMAT(CELL($ws, r, cDate), numfmt="dd.mm.yyyy")
OP         LET(n, n + 1)
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  enq_01.xlsx sheet=EnqSummary cols=A:Sr.Nos.,B:Enquiry,C:Date,D:Item rows=6 data=A:seq(1,1);B:text(8);C:reldates(-2,-1)+boundary(-1,-1,-1,-1,-1,-1);D:values("Head"x3,"Shell"x2,""x1)+boundary("Head","","Shell","Head","Shell","Head")
FIXTURE  enq_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=enq_01.xlsx call=StampEnquiry(EnqSummary) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTIF(EnqSummary![Enquiry], "E-9100") = 5 AND VALUE(EnqSummary![Date]2) > VALUE(EnqSummary![Date]3) AND VALUE(EnqSummary![Date]2) = VALUE(EnqSummary![Date]7)
TEST     fixture=enq_01.xlsx call=StampEnquiry(EnqSummary) assert=VALUE(EnqSummary![Enquiry]3) = PRE_VALUE(EnqSummary![Enquiry]3) AND VALUE(EnqSummary![Date]3) = PRE_VALUE(EnqSummary![Date]3) AND COUNTA(EnqSummary![Item]) = PRE_COUNTA(EnqSummary![Item]) AND COUNTA(EnqSummary![Enquiry]) = PRE_COUNTA(EnqSummary![Enquiry])
TEST     fixture=enq_01.xlsx call=StampEnquiry(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
