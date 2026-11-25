VCL 0.7
TASK
# ASSUME: The sheet is Enq Summary, passed in as a Worksheet; headings are in row 1 (Enquiry, Item), one item per row from row 2
GOAL     Write enquiry number E-9100 in every item row of the Enq Summary sheet.
PROC     StampEnquiry
KIND     Function
MODULE   modEstimation
ARGS     ws As Worksheet
RET      Long
VAR      cEnq As Long
VAR      cItem As Long
VAR      n As Long
PRE      COL($ws, "Enquiry") > 0; COL($ws, "Item") > 0
MODIFIES $ws
OP       LET(cEnq, COL($ws, "Enquiry"))
OP       LET(cItem, COL($ws, "Item"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cItem) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cItem))
OP         SET_VALUE(CELL($ws, r, cEnq), "E-9100")
OP         LET(n, n + 1)
OP       RETURN(n)
ERR      standard
FIXTURE  est_01.xlsx template=estimation
TEST     fixture=est_01.xlsx call=StampEnquiry('Enq Summary') runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE('Enq Summary'!B3) = "E-9100"
TEST     fixture=est_01.xlsx call=StampEnquiry('Enq Summary') assert=RESULT = 2 AND COUNTA('Enq Summary'![Enquiry]) = PRE_COUNTA('Enq Summary'![Enquiry])
TEST     fixture=est_01.xlsx call=StampEnquiry('BOM-RM') assert=RESULT = -1 AND ERROR_COUNT = 1
