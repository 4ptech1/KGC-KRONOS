VCL 0.7
TASK
# ASSUME: the name and the date go in the Prepared By block at the bottom of the WPS sheet: the name in B49 (merged B49:E49) and the date in B50 (merged B50:E50), as the labels Name >> and Date >> show
# ASSUME: today's date is written as a real date with the number format dd/mm/yyyy, so it shows as 26/11/2026 and is a date, not text
# ASSUME: the cells are found by their labels in A49 and A50, which PRE checks, and a re-run writes the same two cells again
GOAL     Write R. Demo, Welding Engineer as the preparer's name and today's date as dd/mm/yyyy in the Prepared By block at the bottom of the WPS sheet.
PROC     SignWps
KIND     Function
MODULE   modWps
ARGS     ws As Worksheet
RET      Long
PRE      CELL($ws, 49, 1) = "Name >>"; CELL($ws, 50, 1) = "Date >>"
MODIFIES $ws
OP       SET_VALUE(CELL($ws, 49, 2), "R. Demo, Welding Engineer")
OP       SET_VALUE(CELL($ws, 50, 2), TODAY)
OP       FORMAT(CELL($ws, 50, 2), numfmt="dd/mm/yyyy")
OP       RETURN(2)
PERF     screen_off
ERR      standard
FIXTURE  wps_01.xlsx template=wps_form
TEST     fixture=wps_01.xlsx call=SignWps(WPS) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(WPS!B49) = "R. Demo, Welding Engineer" AND VALUE(WPS!B50) = TODAY
TEST     fixture=wps_01.xlsx call=SignWps(WPS) assert=VALUE(WPS!A49) = PRE_VALUE(WPS!A49) AND VALUE(WPS!A50) = PRE_VALUE(WPS!A50) AND VALUE(WPS!F49) = PRE_VALUE(WPS!F49) AND VALUE(WPS!C6) = PRE_VALUE(WPS!C6) AND VALUE(WPS!H35) = PRE_VALUE(WPS!H35) AND VALUE(WPS!B48) = PRE_VALUE(WPS!B48) AND ERROR_COUNT = 0
