VCL 0.7
TASK
# ASSUME: The register has one row per item; an enquiry's columns A to I are written once and merged over its item rows.
GOAL     Add a new enquiry with two items to the practice enquiry register and merge its first columns over them.
PROC     AddEnquiry
KIND     Function
MODULE   modRegister
ARGS     ws As Worksheet, enquiry As String, endUser As String
RET      Long
VAR      r As Long
VAR      n As Long
PRE      COL($ws, "Enquiry") = 2; COL($ws, "End User") = 5; COL($ws, "Item") = 10; COL($ws, "Qty") = 15
MODIFIES $ws
OP       EXIT_IF NOT IS_NOTHING(FIND(BLOCK($ws, 2, 2, LASTROW($ws, 10), 2), $enquiry))
OP       LET(r, LASTROW($ws, 10) + 1)
OP       LOOP_ROWS dir=down from=2 to=r - 1 var=i
OP         IF IS_NUMBER(CELL($ws, i, 1))
OP           LET(n, MAX(CELL($ws, i, 1), n))
OP       WRITE_ROW($ws, r, n + 1, $enquiry, "20.11.2026")
OP       SET_VALUE(CELL($ws, r, 5), $endUser)
OP       WRITE_ROW($ws, r, "Suction Bottle", "PV", "Yes", "", "TBA", 2, col=10)
OP       WRITE_ROW($ws, r + 1, "Discharge Bottle", "PV", "Yes", "", "TBA", 2, col=10)
OP       LOOP_COLS dir=right from=1 to=9 var=c
OP         MERGE(BLOCK($ws, r, c, r + 1, c))
OP       RETURN(2)
ERR      standard
FIXTURE  register_02.xlsx template=enquiry_register
TEST     fixture=register_02.xlsx call=AddEnquiry(Sheet1, "E-9104", "Demo Chemicals") runs=2 assert=RESULT = 0 AND VALUE(Sheet1!B8) = "E-9104" AND VALUE(Sheet1!A8) = 4 AND VALUE(Sheet1!J9) = "Discharge Bottle" AND MERGED(Sheet1!A8:A9) = 1 AND MERGED(Sheet1!I8:I9) = 1 AND MERGED(Sheet1!J8:J9) = 0 AND ERROR_COUNT = 0
TEST     fixture=register_02.xlsx call=AddEnquiry(Sheet1, "E-9105", "Demo Gas") assert=RESULT = 2 AND COUNTA(Sheet1!B1:B7) = PRE_COUNTA(Sheet1!B1:B7) AND MERGED(Sheet1!A2:A3) = 1 AND VALUE(Sheet1!E8) = "Demo Gas"
