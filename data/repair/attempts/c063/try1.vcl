VCL 0.6
TASK
# ASSUME: Dispatch headers in row 1 include "Planned Date" and "Actual Date"; Dispatch No is not used
# ASSUME: a text date is text written exactly as yyyy-mm-dd, spaces at the ends allowed; text in any other form, or not a real calendar date, is never guessed and stays as it is
# ASSUME: real dates, blank cells and numbers in these two columns are left alone; each column is read down to its own last value
# ASSUME: the count of converted cells is the Function's result
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
FIXTURE  dispatch_dates_01.xlsx sheet=Dispatch cols=A:Dispatch No,B:Planned Date,C:Actual Date rows=14 data=A:seq(1001,1);B:reldates(-60,0)+boundary("2026-03-05",-40,""," 2026-04-10 ","05/03/2026",-35,"2026-13-45","2026-02-30","2026-05-01","",-20,"TBC",-15,"2026-06-15");C:reldates(-60,0)+boundary("2026-03-09","",-30,"2026-04-12","",-25,"N/A","2026-05-30",-10,"","2026-06-01",-5,"2026-07-04","2026-07-20")
FIXTURE  dispatch_dates_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=dispatch_dates_01.xlsx call=ConvertDispatchTextDates(Dispatch) assert=RESULT = 10 AND ERROR_COUNT = 0 AND VALUE(Dispatch![Planned Date]2) = "2026-03-05" AND VALUE(Dispatch![Planned Date]5) = "2026-04-10" AND VALUE(Dispatch![Planned Date]10) = "2026-05-01" AND VALUE(Dispatch![Planned Date]15) = "2026-06-15" AND VALUE(Dispatch![Actual Date]2) = "2026-03-09" AND VALUE(Dispatch![Actual Date]5) = "2026-04-12" AND VALUE(Dispatch![Actual Date]9) = "2026-05-30" AND VALUE(Dispatch![Actual Date]12) = "2026-06-01" AND VALUE(Dispatch![Actual Date]14) = "2026-07-04" AND VALUE(Dispatch![Actual Date]15) = "2026-07-20"
TEST     fixture=dispatch_dates_01.xlsx call=ConvertDispatchTextDates(Dispatch) assert=VALUE(Dispatch![Planned Date]6) = "05/03/2026" AND VALUE(Dispatch![Planned Date]8) = "2026-13-45" AND VALUE(Dispatch![Planned Date]9) = "2026-02-30" AND VALUE(Dispatch![Planned Date]13) = "TBC" AND VALUE(Dispatch![Actual Date]8) = "N/A" AND COUNTA(Dispatch![Planned Date]) = PRE_COUNTA(Dispatch![Planned Date]) AND COUNTA(Dispatch![Actual Date]) = PRE_COUNTA(Dispatch![Actual Date]) AND COUNTA(Dispatch![Dispatch No]) = PRE_COUNTA(Dispatch![Dispatch No])
TEST     fixture=dispatch_dates_01.xlsx call=ConvertDispatchTextDates(Dispatch) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0
TEST     fixture=dispatch_dates_01.xlsx call=ConvertDispatchTextDates(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1
