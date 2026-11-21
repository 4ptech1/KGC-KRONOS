VCL 0.7
DOCUMENT
# ASSUME: The person typed the items and nozzles in the request; there is no workbook for this enquiry.
GOAL     Fill the practice technical proposal for two bottles from typed values, with the NACE annexure only for sour service.
NAME     FillProposal
TEMPLATE proposal_template.docx
ARGS     proposalNo As String, onDate As String, sour As Boolean
FIELD    Client = "Demo Compressors Ltd"
FIELD    Item_Line = "Pulsation dampeners, 2 types, 4 Nos"
FIELD    End_User = "Example Gas Co"
FIELD    Enquiry_Ref = "E-9101"
FIELD    Proposal_No = $proposalNo
FIELD    Date = $onDate
FIELD    Basis = AI_TEXT("Offer based on the enquiry data sheet and ASME VIII-1, latest edition.")
FIELD    Total_Qty = "4 Nos"
REPEAT   row table=2 from=[["1", "Suction bottle", "DM-V-101", 2], ["2", "Discharge bottle", "DM-V-102", 2]] Sl=1, Item=2, Tag=3, Qty=4
REPEAT   section Equipment from=[["1", "Suction bottle", "DM-V-101", 2, 400, 1600, 10, "18 barg", AI_TEXT("Two gussets on every nozzle.")], ["2", "Discharge bottle", "DM-V-102", 2, 450, 1700, 12, "35 barg", AI_TEXT("Repads as thick as the shell.")]] Sl=1, Item=2, Tag=3, Qty=4, Shell_ID=5, TL_TL=6, Thk=7, Design_Pr=8, Nozzle_Notes=9
REPEAT   row table=4 from=[["DM-V-101", "N1", "Gas inlet", "6 in", "300#"], ["DM-V-101", "N2", "Gas outlet", "6 in", "300#"], ["DM-V-102", "N1", "Gas inlet", "4 in", "600#"]] match=1=Tag Mark=2, Nozzle=3, Size=4, Rating=5
REPEAT   row table=5 from=[["NACE MR0175 / ISO 15156", "Complied"], ["HIC test of plates", "Offered as an option"]] Requirement=1, Confirmation=2
REPEAT   row table=6 from=[["TQ-01", "Suction flanges 150#", "Offered 300# to meet the code"]] TQ_No=1, Requirement=2, Query=3
KEEP     Nace IF $sour
FIXTURE  notes_01.xlsx sheet=Notes cols=A:Note rows=1 data=A:values("practice"x1)
TEST     fixture=notes_01.xlsx call=FillProposal("P-9101-R0", "20 Nov 2026", FALSE) assert=SECTION_COUNT(Equipment) = 2 AND SECTION_COUNT(Nace) = 0 AND TEXT_HAS("NACE") = 0 AND TABLE_ROWS(2) = 4 AND TABLE_ROWS(4) = 3 AND TABLE_CELL(4, 3, 2) = "Gas outlet" AND TABLE_ROWS(6) = 2 AND TABLE_CELL(6, 2, 4) = "600#" AND TABLE_CELL(5, 2, 2) = "Discharge bottle" AND TABLE_CELL(7, 2, 1) = "TQ-01" AND FIELD_TEXT(Proposal_No) = "P-9101-R0" AND LEFT_OVER = 0 AND ERROR_COUNT = 0
TEST     fixture=notes_01.xlsx call=FillProposal("P-9101-R1", "21 Nov 2026", TRUE) assert=SECTION_COUNT(Nace) = 1 AND TEXT_HAS("NACE / Wet H2S Requirements") = 1 AND TABLE_ROWS(7) = 3 AND TABLE_CELL(7, 3, 2) = "Offered as an option" AND TABLE_CELL(8, 2, 1) = "TQ-01" AND TEXT_HAS("Repads as thick as the shell.") = 1
