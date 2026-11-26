VCL 0.7
DOCUMENT
# ASSUME: today's date comes in as the onDate argument
# ASSUME: the request names no equipment, end user, enquiry reference, basis or queries, so those fields are left blank for a person to fill and the equipment and query parts are removed
# ASSUME: the NACE annexure is wanted for this job, so its section is kept once
GOAL     Fill the technical proposal cover for client Demo EPC with proposal P-2750-0002 and add the NACE annexure with its two requirements.
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
TEST     fixture=notes_01.xlsx call=FillNaceProposal("26 Nov 2026") assert=ERROR_COUNT = 0 AND FIELD_TEXT(Client) = "Demo EPC" AND FIELD_TEXT(Proposal_No) = "P-2750-0002" AND FIELD_TEXT(Date) = "26 Nov 2026" AND SECTION_COUNT(Nace) = 1 AND SECTION_COUNT(Equipment) = 0 AND TEXT_HAS("NACE / Wet H2S Requirements") = 1 AND TEXT_HAS("P-2750-0002") = 1 AND LEFT_OVER = 0
