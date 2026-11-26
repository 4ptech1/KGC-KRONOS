VCL 0.7
TASK
# ASSUME: the headings are in row 3 (No. in A, the clarification in B), not row 1, so they are checked by their cells in PRE
# ASSUME: the sheet has no Status column; Status goes in column C with the heading "Status" written in C3
# ASSUME: the new row goes below the last used row of column A and its number is the number in that row plus one
# ASSUME: a last row whose number is not a number stops the run at PRE, as nothing can be numbered on from it
# ASSUME: a run adds one more row each time; the clarification and the status are the words the user gave
GOAL     Add a numbered row at the end of the Assumptions sheet with the clarification Delivery basis 20 weeks Ex-Works and Status Assumed.
PROC     AddDeliveryClarification
KIND     Function
MODULE   modEstimation
ARGS     ws As Worksheet
RET      Long
VAR      last As Long
VAR      num As Long
PRE      CELL($ws, 3, 1) = "No."; CELL($ws, 3, 2) = "Assumption / clarification"; IS_NUMBER(CELL($ws, LASTROW($ws, 1), 1))
MODIFIES $ws
OP       LET(last, LASTROW($ws, 1))
OP       IF IS_NUMBER(CELL($ws, last, 1))
OP         LET(num, CELL($ws, last, 1) + 1)
OP       SET_VALUE(CELL($ws, 3, 3), "Status")
OP       SET_VALUE(CELL($ws, last + 1, 1), num)
OP       SET_VALUE(CELL($ws, last + 1, 2), "Delivery basis 20 weeks Ex-Works")
OP       SET_VALUE(CELL($ws, last + 1, 3), "Assumed")
OP       RETURN(num)
PERF     screen_off
ERR      standard
FIXTURE  est_01.xlsx template=estimation
TEST     fixture=est_01.xlsx call=AddDeliveryClarification('Assumptions & Clarif.') assert=RESULT = 5 AND ERROR_COUNT = 0 AND VALUE('Assumptions & Clarif.'!A8) = 5 AND VALUE('Assumptions & Clarif.'!B8) = "Delivery basis 20 weeks Ex-Works" AND VALUE('Assumptions & Clarif.'!C8) = "Assumed" AND VALUE('Assumptions & Clarif.'!C3) = "Status" AND ROWS_USED('Assumptions & Clarif.') = 8
TEST     fixture=est_01.xlsx call=AddDeliveryClarification('Assumptions & Clarif.') runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND VALUE('Assumptions & Clarif.'!A9) = 6 AND VALUE('Assumptions & Clarif.'!B7) = "Quantities are per the practice enquiry." AND VALUE('Assumptions & Clarif.'!A7) = 4 AND COUNTA('Assumptions & Clarif.'!A:A) = PRE_COUNTA('Assumptions & Clarif.'!A:A) + 2
TEST     fixture=est_01.xlsx call=AddDeliveryClarification(Inputs) assert=RESULT = -1 AND ERROR_COUNT = 1 AND ROWS_USED('Assumptions & Clarif.') = PRE_ROWS_USED('Assumptions & Clarif.')
