VCL 0.6
TASK
# ASSUME: Calibration Tracker headers in row 1 include "Instrument ID", "Next Calibration Due" and "Remark"
# ASSUME: a text due date is real only as day/month/year with slashes (15/03/2026 or 1/4/2026), and the day-first reading is the only one tried
# ASSUME: any other non-blank value that is not a real date, a number included, gets Bad date in Remark; a converted row keeps its Remark
# ASSUME: cells that are blank or hold only spaces are left alone
GOAL     Turn text due dates written day first into real dates in the Calibration Tracker, mark impossible ones Bad date in Remark, and return how many were converted.
PROC     ConvertDueDates
KIND     Function
MODULE   modCalDates
ARGS     ws As Worksheet
RET      Long
VAR      cDue As Long
VAR      cRemark As Long
VAR      parsed As Variant
VAR      converted As Long
PRE      COL($ws, "Next Calibration Due") > 0; COL($ws, "Remark") > 0
MODIFIES $ws["Next Calibration Due", "Remark"]
OP       LET(cDue, COL($ws, "Next Calibration Due"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cDue) var=r
OP         SKIP_IF IS_EMPTY(CELL($ws, r, cDue))
OP         SKIP_IF IS_DATE(CELL($ws, r, cDue))
OP         LET(parsed, PARSE_DATE(CELL($ws, r, cDue), "dd/mm/yyyy"))
OP         IF IS_DATE(parsed)
OP           SET_VALUE(CELL($ws, r, cDue), parsed)
OP           LET(converted, converted + 1)
OP         ELSE
OP           SET_VALUE(CELL($ws, r, cRemark), "Bad date")
OP       RETURN(converted)
PERF     screen_off
ERR      standard
FIXTURE  caltext_01.xlsx sheet=Cal cols=A:Instrument ID,B:Next Calibration Due,C:Remark rows=8 data=A:seq(1,1);B:values("15/03/2026"x1,"31/02/2026"x1,""x1,"1/4/2026"x1,"abc"x1,"29/02/2026"x1,"   "x1,"29/02/2028"x1)+boundary("15/03/2026","31/02/2026","","1/4/2026","abc","29/02/2026","   ","29/02/2028");C:values("note"x8)
FIXTURE  caltext_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  caldates_02.xlsx sheet=Cal cols=A:Instrument ID,B:Next Calibration Due,C:Remark rows=5 data=A:seq(1,1);B:dates(2026-01-01,2026-12-31);C:values("note"x5)
FIXTURE  caldates_02.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=caltext_01.xlsx call=ConvertDueDates(Cal) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(Cal![Remark], "Bad date") = 3 AND COUNTIF(Cal![Remark], "note") = 5 AND COUNTIF(Cal![Next Calibration Due], ">0") = 3
TEST     fixture=caltext_01.xlsx call=ConvertDueDates(Cal) assert=RESULT = 3 AND COUNTIF(Cal![Next Calibration Due], ">0") = PRE_COUNTIF(Cal![Next Calibration Due], ">0") + 3 AND VALUE(Cal![Next Calibration Due]2) = 46096 AND VALUE(Cal![Next Calibration Due]5) = 46113 AND VALUE(Cal![Next Calibration Due]9) = 46812 AND VALUE(Cal![Next Calibration Due]3) = "31/02/2026" AND VALUE(Cal![Remark]3) = "Bad date" AND VALUE(Cal![Remark]4) = "note" AND VALUE(Cal![Remark]8) = "note" AND VALUE(Cal![Next Calibration Due]8) = "   " AND COUNTA(Cal![Instrument ID]) = PRE_COUNTA(Cal![Instrument ID])
TEST     fixture=caldates_02.xlsx call=ConvertDueDates(Cal) assert=RESULT = 0 AND ERROR_COUNT = 0 AND SUM(Cal![Next Calibration Due]) = PRE_SUM(Cal![Next Calibration Due]) AND COUNTIF(Cal![Remark], "note") = 5
TEST     fixture=caltext_01.xlsx call=ConvertDueDates(Notes) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTIF(Cal![Remark], "note") = PRE_COUNTIF(Cal![Remark], "note")
