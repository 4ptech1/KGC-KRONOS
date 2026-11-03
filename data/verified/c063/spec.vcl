VCL 0.6
TASK
# ASSUME: Dispatch headers in row 1 include "Planned Date" and "Actual Date"; Dispatch No is not used
# ASSUME: a text date is text written exactly as yyyy-mm-dd, spaces at the ends allowed; text in any other form, or not a real calendar date, is never guessed and stays as it is
# ASSUME: real dates, blank cells and numbers in these two columns are left alone; each column is read down to its own last value
# ASSUME: the count of converted cells is the Function's result
# ASSUME: the real sheet is "Dispatch"; the TESTs use PlannedText (text dates in Planned Date) and ActualText (text dates in Actual Date), which have the same headings
GOAL     Turn text dates written as yyyy-mm-dd in the Planned Date and Actual Date columns into real dates and report how many cells were converted.
PROC     ConvertDispatchTextDates
KIND     Function
MODULE   modDispatchDates
ARGS     ws As Worksheet
RET      Long
VAR      cPlan As Long
VAR      cActual As Long
VAR      converted As Long
PRE      COL($ws, "Planned Date") > 0; COL($ws, "Actual Date") > 0
MODIFIES $ws["Planned Date", "Actual Date"]
OP       LET(cPlan, COL($ws, "Planned Date"))
OP       LET(cActual, COL($ws, "Actual Date"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cPlan) var=r
OP         IF NOT IS_DATE(CELL($ws, r, cPlan))
OP           IF IS_DATE(PARSE_DATE(CELL($ws, r, cPlan), "yyyy-mm-dd"))
OP             SET_VALUE(CELL($ws, r, cPlan), PARSE_DATE(CELL($ws, r, cPlan), "yyyy-mm-dd"))
OP             LET(converted, converted + 1)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cActual) var=q
OP         IF NOT IS_DATE(CELL($ws, q, cActual))
OP           IF IS_DATE(PARSE_DATE(CELL($ws, q, cActual), "yyyy-mm-dd"))
OP             SET_VALUE(CELL($ws, q, cActual), PARSE_DATE(CELL($ws, q, cActual), "yyyy-mm-dd"))
OP             LET(converted, converted + 1)
OP       RETURN(converted)
PERF     screen_off
ERR      standard
FIXTURE  dispatch_dates_01.xlsx sheet=PlannedText cols=A:Dispatch No,B:Planned Date,C:Actual Date rows=12 data=A:seq(1001,1);B:values("2026-03-05"x1," 2026-04-10 "x1,""x2,"05/03/2026"x1,"2026-13-45"x1,"2026-02-30"x1,"2026-05-01"x1,"TBC"x1,"2026-06-15"x1,"2026-01-31"x1,"2026-07-20"x1)+boundary("2026-03-05"," 2026-04-10 ","","05/03/2026","2026-13-45","2026-02-30","2026-05-01","","TBC","2026-06-15","2026-01-31","2026-07-20");C:dates(2026-01-10,2026-02-25)+blank(2)
FIXTURE  dispatch_dates_01.xlsx sheet=ActualText cols=A:Dispatch No,B:Planned Date,C:Actual Date rows=12 data=A:seq(1001,1);B:dates(2026-01-05,2026-02-20)+blank(2);C:values("2026-03-09"x1,""x3,"2026-04-12"x1,"N/A"x1,"2026-05-30"x1," 2026-06-01 "x1,"2026-00-10"x1,"2026-07-04"x1,"2026-08-15"x1,"2026-07-20"x1)+boundary("2026-03-09","","2026-04-12","N/A","","2026-05-30"," 2026-06-01 ","2026-00-10","2026-07-04","","2026-08-15","2026-07-20")
FIXTURE  dispatch_dates_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=dispatch_dates_01.xlsx call=ConvertDispatchTextDates(PlannedText) assert=RESULT = 6 AND ERROR_COUNT = 0 AND VALUE(PlannedText![Planned Date]2) = "2026-03-05" AND VALUE(PlannedText![Planned Date]3) = "2026-04-10" AND VALUE(PlannedText![Planned Date]8) = "2026-05-01" AND VALUE(PlannedText![Planned Date]11) = "2026-06-15" AND VALUE(PlannedText![Planned Date]12) = "2026-01-31" AND VALUE(PlannedText![Planned Date]13) = "2026-07-20"
TEST     fixture=dispatch_dates_01.xlsx call=ConvertDispatchTextDates(PlannedText) assert=VALUE(PlannedText![Planned Date]5) = "05/03/2026" AND VALUE(PlannedText![Planned Date]6) = "2026-13-45" AND VALUE(PlannedText![Planned Date]7) = "2026-02-30" AND VALUE(PlannedText![Planned Date]10) = "TBC" AND COUNTA(PlannedText![Planned Date]) = PRE_COUNTA(PlannedText![Planned Date]) AND COUNTA(PlannedText![Actual Date]) = PRE_COUNTA(PlannedText![Actual Date]) AND COUNTA(PlannedText![Dispatch No]) = PRE_COUNTA(PlannedText![Dispatch No])
TEST     fixture=dispatch_dates_01.xlsx call=ConvertDispatchTextDates(PlannedText) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0
TEST     fixture=dispatch_dates_01.xlsx call=ConvertDispatchTextDates(ActualText) assert=RESULT = 7 AND ERROR_COUNT = 0 AND VALUE(ActualText![Actual Date]2) = "2026-03-09" AND VALUE(ActualText![Actual Date]4) = "2026-04-12" AND VALUE(ActualText![Actual Date]7) = "2026-05-30" AND VALUE(ActualText![Actual Date]8) = "2026-06-01" AND VALUE(ActualText![Actual Date]10) = "2026-07-04" AND VALUE(ActualText![Actual Date]12) = "2026-08-15" AND VALUE(ActualText![Actual Date]13) = "2026-07-20"
TEST     fixture=dispatch_dates_01.xlsx call=ConvertDispatchTextDates(ActualText) assert=VALUE(ActualText![Actual Date]5) = "N/A" AND VALUE(ActualText![Actual Date]9) = "2026-00-10" AND COUNTA(ActualText![Planned Date]) = PRE_COUNTA(ActualText![Planned Date]) AND COUNTA(ActualText![Actual Date]) = PRE_COUNTA(ActualText![Actual Date]) AND COUNTA(ActualText![Dispatch No]) = PRE_COUNTA(ActualText![Dispatch No])
TEST     fixture=dispatch_dates_01.xlsx call=ConvertDispatchTextDates(ActualText) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0
TEST     fixture=dispatch_dates_01.xlsx call=ConvertDispatchTextDates(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
