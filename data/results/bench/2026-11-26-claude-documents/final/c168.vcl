VCL 0.7
TASK
# ASSUME: The sheet's headers are in row 3 (No. in column A, the clarification in column B); there is no Status header, so Status goes in column C of the new row
# ASSUME: The last numbered row is found from column A; its number is a number, and when it is not, the row count under the headers is used
# ASSUME: Every run adds one more row, as asked, so a re-run is not idempotent
GOAL     Add a numbered Delivery basis clarification row with status Assumed at the end of the assumptions sheet and return its number.
PROC     AddDeliveryClarification
KIND     Function
MODULE   modClarif
ARGS     ws As Worksheet
RET      Long
VAR      last As Long
VAR      n As Long
PRE      SAME(CELL($ws, 3, 1), "No."); SAME(CELL($ws, 3, 2), "Assumption / clarification")
MODIFIES $ws
OP       LET(last, LASTROW($ws, 1))
OP       IF IS_NUMBER(CELL($ws, last, 1))
OP         LET(n, CELL($ws, last, 1) + 1)
OP       ELSE
OP         LET(n, last - 2)
OP       SET_VALUE(CELL($ws, last + 1, 1), n)
OP       SET_VALUE(CELL($ws, last + 1, 2), AI_TEXT("Delivery basis 20 weeks Ex-Works"))
OP       SET_VALUE(CELL($ws, last + 1, 3), AI_TEXT("Assumed"))
OP       RETURN(n)
ERR      standard
FIXTURE  estimate_01.xlsx template=estimation
TEST     fixture=estimate_01.xlsx call=AddDeliveryClarification('Assumptions & Clarif.') runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND VALUE('Assumptions & Clarif.'!A8) = 5 AND VALUE('Assumptions & Clarif.'!A9) = 6 AND VALUE('Assumptions & Clarif.'!B9) = "Delivery basis 20 weeks Ex-Works" AND VALUE('Assumptions & Clarif.'!C9) = "Assumed"
TEST     fixture=estimate_01.xlsx call=AddDeliveryClarification('Assumptions & Clarif.') assert=RESULT = 5 AND ERROR_COUNT = 0 AND COUNTA('Assumptions & Clarif.'!A:A) = PRE_COUNTA('Assumptions & Clarif.'!A:A) + 1 AND VALUE('Assumptions & Clarif.'!B7) = "Quantities are per the practice enquiry."
TEST     fixture=estimate_01.xlsx call=AddDeliveryClarification(Summary) assert=RESULT = -1 AND ERROR_COUNT = 1
