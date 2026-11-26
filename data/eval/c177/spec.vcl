VCL 0.7
DOCUMENT
# ASSUME: the request gives no estimation sheet number and no date, so they are arguments the person types when running it
# ASSUME: no assumptions, weight build-up or cost build-up were given, so those repeated parts are removed with from=[]
# ASSUME: the amounts are typed as the person wrote them, in Indian grouping, and the Sl numbers are text
GOAL     Fill the internal estimation sheet for enquiry E-9105 with the typed client, end user, title, price basis and two priced items with their totals.
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
FIELD    Total_Qty = 4
FIELD    Grand_Total = "11,90,000"
REPEAT   row table=2 from=[["1", "Suction bottle", 2, 410, "2,90,000", "5,80,000"], ["2", "Discharge bottle", 2, 450, "3,05,000", "6,10,000"]] Sl=1, Item=2, Qty=3, Weight=4, Unit_Price=5, Total=6
REPEAT   section Assumption from=[]
REPEAT   column table=3 from=[] Tag=1, Shell_Wt=2, Ends_Wt=3, Fin_Wt=4
REPEAT   section Cost from=[]
FIXTURE  notes_01.xlsx sheet=Notes cols=A:Note rows=1 data=A:values("practice"x1)
TEST     fixture=notes_01.xlsx call=FillEstimateE9105("EST-9105-R0", "26 Nov 2026") assert=FIELD_TEXT(Est_No) = "EST-9105-R0" AND FIELD_TEXT(Date) = "26 Nov 2026" AND FIELD_TEXT(Enquiry) = "E-9105" AND FIELD_TEXT(Title) = "Two pulsation bottles" AND FIELD_TEXT(Client) = "Demo OEM" AND FIELD_TEXT(End_User) = "Example Refinery Ltd" AND FIELD_TEXT(Price_Basis) = "Ex-works, packing included" AND FIELD_TEXT(Total_Qty) = "4" AND FIELD_TEXT(Grand_Total) = "11,90,000" AND TABLE_ROWS(2) = 4 AND TABLE_CELL(2, 2, 1) = "1" AND TABLE_CELL(2, 2, 2) = "Suction bottle" AND TABLE_CELL(2, 2, 3) = "2" AND TABLE_CELL(2, 2, 4) = "410" AND TABLE_CELL(2, 2, 5) = "2,90,000" AND TABLE_CELL(2, 2, 6) = "5,80,000" AND TABLE_CELL(2, 3, 1) = "2" AND TABLE_CELL(2, 3, 2) = "Discharge bottle" AND TABLE_CELL(2, 3, 3) = "2" AND TABLE_CELL(2, 3, 4) = "450" AND TABLE_CELL(2, 3, 5) = "3,05,000" AND TABLE_CELL(2, 3, 6) = "6,10,000" AND TABLE_CELL(2, 4, 3) = "4" AND TABLE_CELL(2, 4, 6) = "11,90,000" AND SECTION_COUNT(Assumption) = 0 AND SECTION_COUNT(Cost) = 0 AND TABLE_COLS(3) = 1 AND LEFT_OVER = 0 AND ERROR_COUNT = 0
TEST     fixture=notes_01.xlsx call=FillEstimateE9105("EST-9105-R1", "27 Nov 2026") assert=TEXT_HAS("Two pulsation bottles") = 1 AND TEXT_HAS("E-9105") = 1 AND TABLE_ROWS(1) = 4 AND TABLE_ROWS(4) = 2 AND TABLE_ROWS(3) = 4
