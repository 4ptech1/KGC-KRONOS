VCL 0.7
TASK
# ASSUME: Headings are in row 1 of the register; each enquiry's first row holds its Sr.Nos., Enquiry, Date, End User and Project, and the item rows below carry only item columns
# ASSUME: The new rows go straight after the last used row of the Item column; Sr.Nos. is one more than the last number in its column
# ASSUME: Date is typed as text 20.11.2026 as in the other rows; only the five asked columns are merged over the two item rows, written through their top-left cell first
# ASSUME: A run that finds E-9104 already in the register does nothing, so a re-run is safe
GOAL     Add enquiry E-9104 with its two items after the last row of the enquiry register and merge its Sr.Nos., Enquiry, Date, End User and Project cells over the two item rows.
PROC     AddEnquiryE9104
KIND     Function
MODULE   modRegister
ARGS     ws As Worksheet
RET      Long
VAR      cSr As Long
VAR      cEnq As Long
VAR      cDate As Long
VAR      cUser As Long
VAR      cProj As Long
VAR      cItem As Long
VAR      cType As Long
VAR      cQty As Long
VAR      last As Long
VAR      sr As Long
PRE      COL($ws, "Sr.Nos.") > 0; COL($ws, "Enquiry") > 0; COL($ws, "Date") > 0; COL($ws, "End User") > 0; COL($ws, "Project") > 0; COL($ws, "Item") > 0; COL($ws, "Type") > 0; COL($ws, "Qty") > 0
MODIFIES $ws
OP       LET(cSr, COL($ws, "Sr.Nos."))
OP       LET(cEnq, COL($ws, "Enquiry"))
OP       LET(cDate, COL($ws, "Date"))
OP       LET(cUser, COL($ws, "End User"))
OP       LET(cProj, COL($ws, "Project"))
OP       LET(cItem, COL($ws, "Item"))
OP       LET(cType, COL($ws, "Type"))
OP       LET(cQty, COL($ws, "Qty"))
OP       EXIT_IF NOT IS_NOTHING(FIND(RANGE($ws, "B2:B200"), "E-9104"))
OP       LET(last, LASTROW($ws, cItem))
OP       LOOP_ROWS dir=up from=last to=2 var=r
OP         EXIT_LOOP_IF IS_NUMBER(CELL($ws, r, cSr))
OP       EXIT_IF NOT IS_NUMBER(CELL($ws, r, cSr))
OP       LET(sr, CELL($ws, r, cSr) + 1)
OP       SET_VALUE(CELL($ws, last + 1, cSr), sr)
OP       SET_VALUE(CELL($ws, last + 1, cEnq), "E-9104")
OP       SET_VALUE(CELL($ws, last + 1, cDate), "20.11.2026")
OP       SET_VALUE(CELL($ws, last + 1, cUser), "Demo Chemicals")
OP       SET_VALUE(CELL($ws, last + 1, cProj), "PRJ-104")
OP       SET_VALUE(CELL($ws, last + 1, cItem), "Knock-out Drum")
OP       SET_VALUE(CELL($ws, last + 1, cType), "PV")
OP       SET_VALUE(CELL($ws, last + 1, cQty), 1)
OP       SET_VALUE(CELL($ws, last + 2, cItem), "Filter Separator")
OP       SET_VALUE(CELL($ws, last + 2, cType), "PV")
OP       SET_VALUE(CELL($ws, last + 2, cQty), 2)
OP       MERGE(BLOCK($ws, last + 1, cSr, last + 2, cSr))
OP       MERGE(BLOCK($ws, last + 1, cEnq, last + 2, cEnq))
OP       MERGE(BLOCK($ws, last + 1, cDate, last + 2, cDate))
OP       MERGE(BLOCK($ws, last + 1, cUser, last + 2, cUser))
OP       MERGE(BLOCK($ws, last + 1, cProj, last + 2, cProj))
OP       RETURN(2)
ERR      standard
FIXTURE  reg_01.xlsx template=enquiry_register
TEST     fixture=reg_01.xlsx call=AddEnquiryE9104(Sheet1) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE(Sheet1![Sr.Nos.]8) = 4 AND VALUE(Sheet1![Enquiry]8) = "E-9104" AND VALUE(Sheet1![Date]8) = "20.11.2026" AND VALUE(Sheet1![End User]8) = "Demo Chemicals" AND VALUE(Sheet1![Project]8) = "PRJ-104" AND VALUE(Sheet1![Item]8) = "Knock-out Drum" AND VALUE(Sheet1![Item]9) = "Filter Separator" AND VALUE(Sheet1![Type]8) = "PV" AND VALUE(Sheet1![Type]9) = "PV" AND VALUE(Sheet1![Qty]8) = 1 AND VALUE(Sheet1![Qty]9) = 2
TEST     fixture=reg_01.xlsx call=AddEnquiryE9104(Sheet1) assert=RESULT = 2 AND MERGED(Sheet1!A8:A9) = 1 AND MERGED(Sheet1!B8:B9) = 1 AND MERGED(Sheet1!C8:C9) = 1 AND MERGED(Sheet1!E8:E9) = 1 AND MERGED(Sheet1!I8:I9) = 1 AND MERGED(Sheet1!D8:D9) = 0 AND MERGED(Sheet1!F8:F9) = 0 AND MERGED(Sheet1!A4:A6) = 1 AND ROWS_USED(Sheet1) = 9 AND VALUE(Sheet1![Enquiry]7) = "E-9103" AND COUNTA(Sheet1![Item]) = PRE_COUNTA(Sheet1![Item]) + 2
