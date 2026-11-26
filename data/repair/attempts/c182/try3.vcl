VCL 0.7
DOCUMENT
# ASSUME: the request gives no header values, items, equipment or basis, so every field is filled with "" and the item rows and equipment sections are removed with from=[]
# ASSUME: the NACE annexure is left out because the job is not sour service, so its section is removed and its table has no rows
# ASSUME: the two queries are typed exactly as the request gives them
GOAL     Write the technical queries table of the proposal with two queries and leave the NACE annexure out of the document.
NAME     FillProposalQueries
TEMPLATE proposal_template.docx
ARGS     book As Workbook
FIELD    Client = ""
FIELD    Item_Line = ""
FIELD    End_User = ""
FIELD    Enquiry_Ref = ""
FIELD    Proposal_No = ""
FIELD    Date = ""
FIELD    Basis = ""
FIELD    Total_Qty = ""
REPEAT   row table=2 from=[] Sl=1, Item=2, Tag=3, Qty=4
REPEAT   section Equipment from=[]
REPEAT   row table=5 from=[] Requirement=1, Confirmation=2
REPEAT   row table=6 from=[["TQ-01", "Suction flanges 150#", "Offered 300# to meet the code"], ["TQ-02", "Inspection opening rating blank", "Offered 600#"]] TQ_No=1, Requirement=2, Query=3
KEEP     Nace IF FALSE
FIXTURE  notes_01.xlsx sheet=Notes cols=A:Note rows=1 data=A:values("practice"x1)
TEST     fixture=notes_01.xlsx call=FillProposalQueries(notes_01.xlsx) assert=SECTION_COUNT(Nace) = 0 AND SECTION_COUNT(Equipment) = 0 AND TEXT_HAS("NACE") = 0 AND TABLE_ROWS(3) = 3 AND TABLE_CELL(3, 2, 1) = "TQ-01" AND TABLE_CELL(3, 2, 2) = "Suction flanges 150#" AND TABLE_CELL(3, 2, 3) = "Offered 300# to meet the code" AND TABLE_CELL(3, 3, 1) = "TQ-02" AND TABLE_CELL(3, 3, 2) = "Inspection opening rating blank" AND TABLE_CELL(3, 3, 3) = "Offered 600#" AND LEFT_OVER = 0 AND ERROR_COUNT = 0
TEST     fixture=notes_01.xlsx call=FillProposalQueries(notes_01.xlsx) assert=TABLE_ROWS(2) = 2 AND TABLE_ROWS(1) = 6 AND FIELD_TEXT(Client) = "" AND FIELD_TEXT(Total_Qty) = "" AND FIELD_TEXT(Enquiry_Ref) = "" AND TEXT_HAS("Technical Queries") = 1 AND TEXT_HAS("Technical Delivery Conditions") = 1
