VCL 0.7
TASK
# ASSUME: The enquiry number is in the Enquiry column, written in the first row of the enquiry; its Result cell is merged over the item rows and is written in that first row only
# ASSUME: Headings are in row 1; a missing enquiry changes nothing and the macro returns 0
GOAL     Set the Result of enquiry E-9102 to Won in the first cell of its merged Result range in the enquiry register.
PROC     SetEnquiryResultWon
KIND     Function
MODULE   modRegister
ARGS     ws As Worksheet
RET      Long
VAR      cEnq As Long
VAR      cRes As Long
VAR      hit As Range
PRE      COL($ws, "Enquiry") > 0; COL($ws, "Result") > 0
MODIFIES $ws["Result"]
OP       LET(cEnq, COL($ws, "Enquiry"))
OP       LET(cRes, COL($ws, "Result"))
OP       LET(hit, FIND(RANGE($ws, "B2:B200"), "E-9102"))
OP       IF IS_NOTHING(hit)
OP         RETURN(0)
OP       SET_VALUE(CELL($ws, ROW_OF(hit), cRes), "Won")
OP       RETURN(1)
ERR      standard
FIXTURE  reg_01.xlsx template=enquiry_register
TEST     fixture=reg_01.xlsx call=SetEnquiryResultWon(Sheet1) runs=2 assert=RESULT = 1 AND ERROR_COUNT = 0 AND VALUE(Sheet1![Result]4) = "Won" AND MERGED(Sheet1!D4:D6) = 1
TEST     fixture=reg_01.xlsx call=SetEnquiryResultWon(Sheet1) assert=VALUE(Sheet1![Result]2) = "Lost" AND COUNTA(Sheet1![Result]) = PRE_COUNTA(Sheet1![Result]) + 1 AND COUNTIF(Sheet1![Result], "Won") = 1 AND VALUE(Sheet1![Enquiry]4) = "E-9102" AND VALUE(Sheet1![Result]7) = ""
