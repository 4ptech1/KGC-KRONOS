VCL 0.7
DOCUMENT
# ASSUME: the estimation workbook is passed as a Workbook argument; its Assumptions & Clarif. sheet has the numbers in column A and the clarification in column B, from row 4 to the last used row of column A
# ASSUME: one paragraph per row, taking the text of column B only
# ASSUME: the request gives no header values, price rows, weights or costs, so every field is filled with "" and the price rows, the weight columns and the cost sections are removed with from=[]
GOAL     Write the basis and assumptions section of the internal estimation sheet with one paragraph for each row of the workbook's assumptions sheet.
NAME     FillEstimateAssumptions
TEMPLATE estimate_template.docx
ARGS     book As Workbook
FIELD    Est_No = ""
FIELD    Date = ""
FIELD    Enquiry = ""
FIELD    Title = ""
FIELD    Client = ""
FIELD    End_User = ""
FIELD    Price_Basis = ""
FIELD    Total_Qty = ""
FIELD    Grand_Total = ""
REPEAT   row table=2 from=[] Sl=1, Item=2, Qty=3, Weight=4, Unit_Price=5, Total=6
REPEAT   section Assumption from=BLOCK(SHEET($book, "Assumptions & Clarif."), 4, 2, LASTROW(SHEET($book, "Assumptions & Clarif."), 2), 2) Text=1
REPEAT   column table=3 from=[] Tag=1, Shell_Wt=2, Ends_Wt=3, Fin_Wt=4
REPEAT   section Cost from=[]
FIXTURE  est_01.xlsx template=estimation
TEST     fixture=est_01.xlsx call=FillEstimateAssumptions(est_01.xlsx) assert=SECTION_COUNT(Assumption) = 4 AND TEXT_HAS("Prices ex-works, packing included; freight and taxes extra.") = 1 AND TEXT_HAS("Rates are invented practice rates.") = 1 AND TEXT_HAS("Option 1 has no shop hydrotest; option 2 uses a test cover and ring.") = 1 AND TEXT_HAS("Quantities are per the practice enquiry.") = 1 AND SECTION_COUNT(Cost) = 0 AND TABLE_ROWS(2) = 2 AND TABLE_COLS(3) = 1 AND LEFT_OVER = 0 AND ERROR_COUNT = 0
TEST     fixture=est_01.xlsx call=FillEstimateAssumptions(est_01.xlsx) assert=TEXT_HAS("Assumptions & Clarifications") = 0 AND TEXT_HAS("Option 1 adder") = 0 AND FIELD_TEXT(Est_No) = "" AND FIELD_TEXT(Enquiry) = "" AND FIELD_TEXT(Client) = "" AND FIELD_TEXT(Price_Basis) = "" AND TABLE_ROWS(1) = 4 AND TABLE_CELL(2, 2, 2) = "Total (excl. GST)"
