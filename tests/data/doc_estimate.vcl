VCL 0.7
DOCUMENT
# ASSUME: The estimation workbook was saved in Excel after its last change, so every formula holds its value.
GOAL     Fill the practice internal estimation sheet from the practice estimation workbook.
NAME     FillEstimateSheet
TEMPLATE estimate_template.docx
ARGS     wb As Workbook, estNo As String, onDate As String
FIELD    Est_No = $estNo
FIELD    Date = $onDate
FIELD    Enquiry = CELL(SHEET($wb, "Enq Summary"), 2, 2)
FIELD    Title = "Channel and floating head assemblies, 2 sets"
FIELD    Client = CELL(SHEET($wb, "Enq Summary"), 2, 8)
FIELD    End_User = CELL(SHEET($wb, "Enq Summary"), 2, 5)
FIELD    Price_Basis = AI_TEXT("Ex-works, packing included; freight and taxes extra.")
FIELD    Total_Qty = CELL(SHEET($wb, "Enq Summary"), 2, 15) + CELL(SHEET($wb, "Enq Summary"), 3, 15)
FIELD    Grand_Total = NUMBER_TEXT(CELL(SHEET($wb, "Summary"), 17, 4), "#,##,##0")
REPEAT   row table=2 from=[["1", CELL(SHEET($wb, "Enq Summary"), 2, 10), 2, NUMBER_TEXT(CELL(SHEET($wb, "Enq Summary"), 2, 21), "0"), NUMBER_TEXT(CELL(SHEET($wb, "Summary"), 17, 2), "#,##,##0"), NUMBER_TEXT(2 * CELL(SHEET($wb, "Summary"), 17, 2), "#,##,##0")], ["2", CELL(SHEET($wb, "Enq Summary"), 3, 10), 2, NUMBER_TEXT(CELL(SHEET($wb, "Enq Summary"), 3, 21), "0"), NUMBER_TEXT(CELL(SHEET($wb, "Summary"), 17, 3), "#,##,##0"), NUMBER_TEXT(2 * CELL(SHEET($wb, "Summary"), 17, 3), "#,##,##0")]] Sl=1, Item=2, Qty=3, Weight=4, Unit_Price=5, Total=6
REPEAT   section Assumption from=RANGE(SHEET($wb, "Assumptions & Clarif."), "B4:B7") Text=1
REPEAT   column table=3 from=[[CELL(SHEET($wb, "Enq Summary"), 2, 13), NUMBER_TEXT(CELL(SHEET($wb, "BOM-RM"), 6, 12) / 2, "0"), NUMBER_TEXT(CELL(SHEET($wb, "BOM-RM"), 5, 12) / 2, "0"), NUMBER_TEXT(CELL(SHEET($wb, "Enq Summary"), 2, 21), "0")], [CELL(SHEET($wb, "Enq Summary"), 3, 13), NUMBER_TEXT(CELL(SHEET($wb, "BOM-RM"), 8, 12) / 2, "0"), NUMBER_TEXT(CELL(SHEET($wb, "BOM-RM"), 7, 12) / 2, "0"), NUMBER_TEXT(CELL(SHEET($wb, "Enq Summary"), 3, 21), "0")]] Tag=1, Shell_Wt=2, Ends_Wt=3, Fin_Wt=4
REPEAT   section Cost from=[["1", CELL(SHEET($wb, "Enq Summary"), 2, 10), NUMBER_TEXT(CELL(SHEET($wb, "Summary"), 5, 2) / 2, "#,##,##0"), NUMBER_TEXT((CELL(SHEET($wb, "Summary"), 6, 2) + CELL(SHEET($wb, "Summary"), 7, 2)) / 2, "#,##,##0"), NUMBER_TEXT(CELL(SHEET($wb, "Summary"), 8, 2) / 2, "#,##,##0"), NUMBER_TEXT(CELL(SHEET($wb, "Summary"), 17, 2), "#,##,##0")], ["2", CELL(SHEET($wb, "Enq Summary"), 3, 10), NUMBER_TEXT(CELL(SHEET($wb, "Summary"), 5, 3) / 2, "#,##,##0"), NUMBER_TEXT((CELL(SHEET($wb, "Summary"), 6, 3) + CELL(SHEET($wb, "Summary"), 7, 3)) / 2, "#,##,##0"), NUMBER_TEXT(CELL(SHEET($wb, "Summary"), 8, 3) / 2, "#,##,##0"), NUMBER_TEXT(CELL(SHEET($wb, "Summary"), 17, 3), "#,##,##0")]] Sl=1, Item=2, Materials=3, Conversion=4, Factory_Cost=5, Unit_Price=6
FIXTURE  est_01.xlsx template=estimation
TEST     fixture=est_01.xlsx call=FillEstimateSheet(est_01.xlsx, "EST-9001", "20 Nov 2026") assert=FIELD_TEXT(Grand_Total) = "6,76,000" AND FIELD_TEXT(Enquiry) = "E-9100 (TBA)" AND FIELD_TEXT(Est_No) = "EST-9001" AND TABLE_ROWS(2) = 4 AND TABLE_CELL(2, 2, 2) = "Channel Head Assembly" AND TABLE_CELL(2, 3, 5) = "1,53,000" AND TABLE_CELL(2, 3, 6) = "3,06,000" AND TABLE_COLS(3) = 3 AND TABLE_CELL(3, 1, 3) = "DM-1001B" AND SECTION_COUNT(Assumption) = 4 AND SECTION_COUNT(Cost) = 2 AND TEXT_HAS("Rates are invented practice rates.") = 1 AND LEFT_OVER = 0 AND ERROR_COUNT = 0
TEST     fixture=est_01.xlsx call=FillEstimateSheet(est_01.xlsx, "EST-9002", "21 Nov 2026") assert=FIELD_TEXT(Date) = "21 Nov 2026" AND FIELD_TEXT(Total_Qty) = "4" AND TABLE_CELL(4, 5, 2) = "1,85,000" AND TABLE_ROWS(6) = 2
