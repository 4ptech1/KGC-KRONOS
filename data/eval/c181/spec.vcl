VCL 0.7
DOCUMENT
# ASSUME: Today is the date KRONOS gives the run; the recipe takes it as the argument onDate, typed as text such as 25 Nov 2026.
# ASSUME: The request gives no item line, end user, enquiry ref, basis, items, equipment or queries, so those fields are "" and those parts are removed.
# ASSUME: The NACE annexure is kept because the job is sour service.
GOAL     Fill the proposal cover with the client and proposal number and keep the NACE annexure with its two requirements.
NAME     FillNaceProposal
TEMPLATE proposal_template.docx
ARGS     onDate As String
FIELD    Client = "Demo EPC"
FIELD    Item_Line = ""
FIELD    End_User = ""
FIELD    Enquiry_Ref = ""
FIELD    Proposal_No = "P-2750-0002"
FIELD    Date = $onDate
FIELD    Basis = ""
FIELD    Total_Qty = ""
REPEAT   row table=2 from=[] Sl=1, Item=2, Tag=3, Qty=4
REPEAT   section Equipment from=[]
REPEAT   row table=5 from=[["NACE MR0175 / ISO 15156", "Complied"], ["HIC test of plates", "Offered as an option"]] Requirement=1, Confirmation=2
REPEAT   row table=6 from=[] TQ_No=1, Requirement=2, Query=3
KEEP     Nace IF TRUE
FIXTURE  notes_01.xlsx sheet=Notes cols=A:Note rows=1 data=A:values("practice"x1)
TEST     fixture=notes_01.xlsx call=FillNaceProposal("25 Nov 2026") assert=SECTION_COUNT(Nace) = 1 AND SECTION_COUNT(Equipment) = 0 AND FIELD_TEXT(Client) = "Demo EPC" AND FIELD_TEXT(Proposal_No) = "P-2750-0002" AND FIELD_TEXT(Date) = "25 Nov 2026" AND TEXT_HAS("NACE / Wet H2S Requirements") = 1 AND TABLE_ROWS(3) = 3 AND TABLE_CELL(3, 2, 1) = "NACE MR0175 / ISO 15156" AND TABLE_CELL(3, 2, 2) = "Complied" AND TABLE_CELL(3, 3, 1) = "HIC test of plates" AND TABLE_CELL(3, 3, 2) = "Offered as an option" AND LEFT_OVER = 0 AND ERROR_COUNT = 0
TEST     fixture=notes_01.xlsx call=FillNaceProposal("26 Nov 2026") assert=FIELD_TEXT(Date) = "26 Nov 2026" AND TABLE_CELL(1, 6, 3) = "26 Nov 2026" AND TABLE_CELL(1, 1, 3) = "Demo EPC" AND TABLE_ROWS(2) = 2 AND TABLE_ROWS(4) = 1
