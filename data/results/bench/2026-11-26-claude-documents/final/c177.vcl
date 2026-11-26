VCL 0.7
DOCUMENT
# ASSUME: the request gives no per-item weight build-up or cost build-up, so the weight table (table 3) and the Cost section are removed rather than filled with guesses
# ASSUME: the Basis and Assumptions section gets one remark written by the AI, shown for a person to check
# ASSUME: the quantity, weight, price and total figures are typed text exactly as given, with Indian digit grouping
GOAL     Fill the internal estimation sheet for enquiry E-9105 with its client, end user, price basis and the two bottle items with their totals.
NAME     FillEstimateE9105
TEMPLATE estimate_template.docx
ARGS     estNo As String, onDate As String
FIELD    Est_No = $estNo
FIELD    Date = $onDate
FIELD    Enquiry = "E-9105"
FIELD    Title = "Two pulsation bottles"
FIELD    Client = "Demo OEM"
FIELD    End_User = "Example Refinery Ltd"
FIELD    Price_Basis = "Ex-works, packing included"
FIELD    Total_Qty = "4"
FIELD    Grand_Total = "11,90,000"
REPEAT   row table=2 from=[["1", "Suction bottle", "2", "410", "2,90,000", "5,80,000"], ["2", "Discharge bottle", "2", "450", "3,05,000", "6,10,000"]] Sl=1, Item=2, Qty=3, Weight=4, Unit_Price=5, Total=6
REPEAT   section Assumption from=[[AI_TEXT("Prices are as typed in the request; no weight or cost build-up was supplied.")]] Text=1
REPEAT   column table=3 from=[] Tag=1, Shell_Wt=2, Ends_Wt=3, Fin_Wt=4
REPEAT   section Cost from=[]
FIXTURE  notes_01.xlsx sheet=Notes cols=A:Note rows=1 data=A:values("practice"x1)
TEST     fixture=notes_01.xlsx call=FillEstimateE9105("EST-9105", "26 Nov 2026") assert=ERROR_COUNT = 0 AND FIELD_TEXT(Enquiry) = "E-9105" AND FIELD_TEXT(Client) = "Demo OEM" AND FIELD_TEXT(End_User) = "Example Refinery Ltd" AND FIELD_TEXT(Price_Basis) = "Ex-works, packing included" AND TABLE_ROWS(2) = 4 AND TABLE_CELL(2, 2, 2) = "Suction bottle" AND TABLE_CELL(2, 3, 5) = "3,05,000" AND TABLE_CELL(2, 4, 3) = "4" AND TABLE_CELL(2, 4, 6) = "11,90,000" AND SECTION_COUNT(Assumption) = 1 AND SECTION_COUNT(Cost) = 0 AND TEXT_HAS("EST-9105") = 1 AND TEXT_HAS("26 Nov 2026") = 1 AND LEFT_OVER = 0
