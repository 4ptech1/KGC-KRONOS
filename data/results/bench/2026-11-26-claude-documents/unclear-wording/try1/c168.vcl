VCL 0.7
TASK
# ASSUME: the sheet has its title in row 1, headers in row 3 (No. in A, Assumption / clarification in B) and the numbered rows below
# ASSUME: the sheet has no Status column, so Status goes in column C, with the header Status written to C3 when C3 is blank
# ASSUME: the new number is the last number in column A plus one; a last number that is not a number gives 1
# ASSUME: a re-run adds nothing when that clarification is already on the sheet; the function returns the new number, or 0 when nothing was added
GOAL     Add the clarification Delivery basis 20 weeks Ex-Works with Status Assumed as the next numbered row of the Assumptions and Clarifications sheet.
PROC     AddDeliveryClarification
KIND     Function
MODULE   modEstimate
ARGS     ws As Worksheet
RET      Long
VAR      last As Long
VAR      num As Long
VAR      found As Boolean
PRE      SAME(CELL($ws, 3, 1), "No."); SAME(CELL($ws, 3, 2), "Assumption / clarification")
MODIFIES $ws
OP       LET(last, LASTROW($ws, 1))
OP       LOOP_ROWS dir=down from=4 to=last var=r
OP         IF SAME(CELL($ws, r, 2), "Delivery basis 20 weeks Ex-Works") THEN LET(found, TRUE)
OP       EXIT_IF found
OP       LET(num, 1)
OP       IF last >= 4
OP         IF IS_NUMBER(CELL($ws, last, 1))
OP           LET(num, CELL($ws, last, 1) + 1)
OP       IF IS_EMPTY(CELL($ws, 3, 3)) THEN SET_VALUE(CELL($ws, 3, 3), "Status")
OP       SET_VALUE(CELL($ws, last + 1, 1), num)
OP       SET_VALUE(CELL($ws, last + 1, 2), "Delivery basis 20 weeks Ex-Works")
OP       SET_VALUE(CELL($ws, last + 1, 3), "Assumed")
OP       RETURN(num)
ERR      standard
FIXTURE  estimate_01.xlsx template=estimation
TEST     fixture=estimate_01.xlsx call=AddDeliveryClarification('Assumptions & Clarif.') runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTA('Assumptions & Clarif.'![B]) = 6
TEST     fixture=estimate_01.xlsx call=AddDeliveryClarification('Assumptions & Clarif.') assert=RESULT = 5 AND ERROR_COUNT = 0 AND VALUE('Assumptions & Clarif.'!A8) = 5 AND VALUE('Assumptions & Clarif.'!B8) = "Delivery basis 20 weeks Ex-Works" AND VALUE('Assumptions & Clarif.'!C8) = "Assumed" AND VALUE('Assumptions & Clarif.'!B7) = "Quantities are per the practice enquiry." AND ROWS_USED('Assumptions & Clarif.') = 8
TEST     fixture=estimate_01.xlsx call=AddDeliveryClarification(Summary) assert=RESULT = -1 AND ERROR_COUNT = 1
