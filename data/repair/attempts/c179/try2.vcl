VCL 0.7
DOCUMENT
# ASSUME: The request gives no items, equipment, nozzles, NACE requirements or queries, so those parts are removed and Total_Qty is filled with "".
# ASSUME: The basis of offer is a draft sentence for the engineer to check, so it is AI_TEXT.
GOAL     Fill the proposal's cover and header with the client, item line, end user, references and a draft basis of offer.
NAME     FillProposalCover
TEMPLATE proposal_template.docx
FIELD    Client = "Demo EPC"
FIELD    Item_Line = "Pulsation bottles, 4 Nos"
FIELD    End_User = "Example Gas Co"
FIELD    Enquiry_Ref = "E-9101"
FIELD    Proposal_No = "P-2750-0001"
FIELD    Date = "25 Nov 2026"
FIELD    Basis = AI_TEXT("Offer is based on the enquiry data sheets and ASME VIII Div 1, latest edition.")
FIELD    Total_Qty = ""
REPEAT   row table=2 from=[] Sl=1, Item=2, Tag=3, Qty=4
REPEAT   section Equipment from=[]
REPEAT   row table=5 from=[] Requirement=1, Confirmation=2
REPEAT   row table=6 from=[] TQ_No=1, Requirement=2, Query=3
KEEP     Nace IF FALSE
FIXTURE  notes_01.xlsx sheet=Notes cols=A:Note rows=1 data=A:values("practice"x1)
TEST     fixture=notes_01.xlsx call=FillProposalCover() assert=FIELD_TEXT(Client) = "Demo EPC" AND FIELD_TEXT(Proposal_No) = "P-2750-0001" AND FIELD_TEXT(Date) = "25 Nov 2026" AND FIELD_TEXT(Item_Line) = "Pulsation bottles, 4 Nos" AND FIELD_TEXT(End_User) = "Example Gas Co" AND FIELD_TEXT(Enquiry_Ref) = "E-9101" AND TEXT_HAS("Offer is based on the enquiry data sheets") = 1 AND SECTION_COUNT(Equipment) = 0 AND SECTION_COUNT(Nace) = 0 AND LEFT_OVER = 0 AND ERROR_COUNT = 0
TEST     fixture=notes_01.xlsx call=FillProposalCover() assert=TABLE_ROWS(1) = 6 AND TABLE_CELL(1, 1, 3) = "Demo EPC" AND TABLE_CELL(1, 5, 3) = "P-2750-0001" AND TABLE_CELL(1, 6, 3) = "25 Nov 2026" AND TABLE_ROWS(2) = 2 AND TABLE_ROWS(3) = 1 AND TEXT_HAS("NACE") = 0
