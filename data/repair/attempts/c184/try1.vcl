VCL 0.7
DOCUMENT
# ASSUME: Today's date is the argument onDate, typed as text such as 25 Nov 2026.
# ASSUME: Basis of offer, items table, equipment sections, NACE annexure and queries are not asked for, so Basis and Total_Qty are "" and those parts are removed.
GOAL     Fill the proposal cover for Demo Chemicals with the enquiry, proposal number, date and item line, with no equipment or NACE sections.
NAME     FillCoverOnly
TEMPLATE proposal_template.docx
ARGS     onDate As String
FIELD    Client = "Demo Chemicals"
FIELD    Item_Line = "Knock-out drum and filter separator, 3 Nos"
FIELD    End_User = "Demo Chemicals"
FIELD    Enquiry_Ref = "E-9104"
FIELD    Proposal_No = "P-2750-0003"
FIELD    Date = $onDate
FIELD    Basis = ""
FIELD    Total_Qty = ""
REPEAT   row table=2 from=[] Sl=1, Item=2, Tag=3, Qty=4
REPEAT   section Equipment from=[]
REPEAT   row table=5 from=[] Requirement=1, Confirmation=2
REPEAT   row table=6 from=[] TQ_No=1, Requirement=2, Query=3
KEEP     Nace IF FALSE
FIXTURE  notes_01.xlsx sheet=Notes cols=A:Note rows=1 data=A:values("practice"x1)
TEST     fixture=notes_01.xlsx call=FillCoverOnly("25 Nov 2026") assert=FIELD_TEXT(Client) = "Demo Chemicals" AND FIELD_TEXT(End_User) = "Demo Chemicals" AND FIELD_TEXT(Item_Line) = "Knock-out drum and filter separator, 3 Nos" AND FIELD_TEXT(Enquiry_Ref) = "E-9104" AND FIELD_TEXT(Proposal_No) = "P-2750-0003" AND FIELD_TEXT(Date) = "25 Nov 2026" AND SECTION_COUNT(Equipment) = 0 AND SECTION_COUNT(Nace) = 0 AND TEXT_HAS("NACE") = 0 AND LEFT_OVER = 0 AND ERROR_COUNT = 0
TEST     fixture=notes_01.xlsx call=FillCoverOnly("26 Nov 2026") assert=TABLE_ROWS(1) = 6 AND TABLE_CELL(1, 2, 3) = "Knock-out drum and filter separator, 3 Nos" AND TABLE_CELL(1, 6, 3) = "26 Nov 2026" AND TABLE_CELL(1, 4, 3) = "E-9104" AND TABLE_ROWS(2) = 2 AND TABLE_ROWS(3) = 1
