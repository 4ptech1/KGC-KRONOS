VCL 0.7
DOCUMENT
# ASSUME: the request gives no client, end user, item line, proposal number, date or basis, so those fields are filled with ""
# ASSUME: the total quantity is the two quantities added, 4 Nos
# ASSUME: the design pressures are written as typed with the unit barg, the nozzle sizes with in, and no nozzle notes were given, so Nozzle_Notes is ""
# ASSUME: the request asks for no NACE annexure and no technical queries, so the NACE section is removed and the queries table has no rows
GOAL     Make the technical proposal for enquiry E-9101 with one equipment section for the suction bottle and one for the discharge bottle, each with its design data and nozzles.
NAME     FillProposalE9101
TEMPLATE proposal_template.docx
FIELD    Client = ""
FIELD    Item_Line = ""
FIELD    End_User = ""
FIELD    Enquiry_Ref = "E-9101"
FIELD    Proposal_No = ""
FIELD    Date = ""
FIELD    Basis = ""
FIELD    Total_Qty = "4 Nos"
REPEAT   row table=2 from=[["1", "Suction bottle", "DM-V-101", 2], ["2", "Discharge bottle", "DM-V-102", 2]] Sl=1, Item=2, Tag=3, Qty=4
REPEAT   section Equipment from=[["1", "Suction bottle", "DM-V-101", 2, 400, 1600, 10, "18 barg", ""], ["2", "Discharge bottle", "DM-V-102", 2, 450, 1800, 10, "35 barg", ""]] Sl=1, Item=2, Tag=3, Qty=4, Shell_ID=5, TL_TL=6, Thk=7, Design_Pr=8, Nozzle_Notes=9
REPEAT   row table=4 from=[["DM-V-101", "N1", "Gas inlet", "6 in", "300#"], ["DM-V-101", "N2", "Gas outlet", "6 in", "300#"], ["DM-V-101", "N3", "Drain", "2 in", "300#"], ["DM-V-102", "N1", "Gas inlet", "4 in", "600#"], ["DM-V-102", "N2", "Gas outlet", "4 in", "600#"], ["DM-V-102", "N3", "Drain", "1 in", "600#"]] match=1=Tag Mark=2, Nozzle=3, Size=4, Rating=5
REPEAT   row table=5 from=[] Requirement=1, Confirmation=2
REPEAT   row table=6 from=[] TQ_No=1, Requirement=2, Query=3
KEEP     Nace IF FALSE
FIXTURE  notes_01.xlsx sheet=Notes cols=A:Note rows=1 data=A:values("practice"x1)
TEST     fixture=notes_01.xlsx call=FillProposalE9101() assert=SECTION_COUNT(Equipment) = 2 AND SECTION_COUNT(Nace) = 0 AND FIELD_TEXT(Enquiry_Ref) = "E-9101" AND FIELD_TEXT(Total_Qty) = "4 Nos" AND TABLE_ROWS(2) = 4 AND TABLE_CELL(2, 2, 2) = "Suction bottle" AND TABLE_CELL(2, 2, 3) = "DM-V-101" AND TABLE_CELL(2, 2, 4) = "2" AND TABLE_CELL(2, 3, 2) = "Discharge bottle" AND TABLE_CELL(2, 3, 3) = "DM-V-102" AND TABLE_CELL(2, 3, 4) = "2" AND TABLE_CELL(3, 2, 2) = "Suction bottle" AND TABLE_CELL(3, 3, 2) = "DM-V-101 / 2 Nos" AND TABLE_CELL(3, 4, 2) = "400 ID x 1600 TL-TL x 10 thk" AND TABLE_CELL(3, 5, 2) = "18 barg" AND TABLE_ROWS(4) = 4 AND TABLE_CELL(4, 2, 1) = "N1" AND TABLE_CELL(4, 2, 2) = "Gas inlet" AND TABLE_CELL(4, 2, 3) = "6 in" AND TABLE_CELL(4, 2, 4) = "300#" AND TABLE_CELL(4, 3, 2) = "Gas outlet" AND TABLE_CELL(4, 4, 1) = "N3" AND TABLE_CELL(4, 4, 2) = "Drain" AND TABLE_CELL(4, 4, 3) = "2 in" AND LEFT_OVER = 0 AND ERROR_COUNT = 0
TEST     fixture=notes_01.xlsx call=FillProposalE9101() assert=TABLE_CELL(5, 2, 2) = "Discharge bottle" AND TABLE_CELL(5, 3, 2) = "DM-V-102 / 2 Nos" AND TABLE_CELL(5, 4, 2) = "450 ID x 1800 TL-TL x 10 thk" AND TABLE_CELL(5, 5, 2) = "35 barg" AND TABLE_ROWS(6) = 4 AND TABLE_CELL(6, 2, 1) = "N1" AND TABLE_CELL(6, 2, 3) = "4 in" AND TABLE_CELL(6, 2, 4) = "600#" AND TABLE_CELL(6, 3, 2) = "Gas outlet" AND TABLE_CELL(6, 4, 2) = "Drain" AND TABLE_CELL(6, 4, 3) = "1 in" AND TABLE_CELL(6, 4, 4) = "600#" AND TABLE_ROWS(7) = 1 AND TEXT_HAS("NACE") = 0 AND TEXT_HAS("Nozzle notes: ") = 1
