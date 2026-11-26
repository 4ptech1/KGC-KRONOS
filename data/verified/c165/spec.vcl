VCL 0.7
TASK
# ASSUME: the name and the date go in the Prepared By block at the bottom of the WPS sheet, in the cell to the right of the labels Name >> and Date >> (B49 and B50 on the practice form, merged to E)
# ASSUME: the labels are found with FIND in column A, so a form with the block on other rows still works; a sheet without both labels is stopped by EXIT_IF and nothing is written
# ASSUME: today's date is written as a real date with the number format dd/mm/yyyy, so it shows as 26/11/2026
# ASSUME: the practice sheet is a plain sheet named WPS with Label, Value and Ref columns, not the merged form, so TESTs can compare the date with a reldates column
GOAL     Write R. Demo, Welding Engineer as the preparer's name and today's date as dd/mm/yyyy beside the Name and Date labels at the bottom of the WPS sheet.
PROC     SignWps
KIND     Function
MODULE   modWps
ARGS     ws As Worksheet
RET      Long
VAR      nameHit As Range
VAR      dateHit As Range
MODIFIES $ws
OP       LET(nameHit, FIND(USED_RANGE($ws), "Name >>", match=whole))
OP       LET(dateHit, FIND(USED_RANGE($ws), "Date >>", match=whole))
OP       EXIT_IF IS_NOTHING(nameHit)
OP       EXIT_IF IS_NOTHING(dateHit)
OP       SET_VALUE(CELL($ws, ROW_OF(nameHit), 2), "R. Demo, Welding Engineer")
OP       SET_VALUE(CELL($ws, ROW_OF(dateHit), 2), TODAY)
OP       FORMAT(CELL($ws, ROW_OF(dateHit), 2), numfmt="dd/mm/yyyy")
OP       RETURN(2)
PERF     screen_off
ERR      standard
FIXTURE  wps_01.xlsx sheet=WPS cols=A:Label,B:Value,C:Ref rows=4 data=A:values("Other"x1,"Name >>"x1,"Date >>"x1,"Remarks"x1)+boundary("Other","Name >>","Date >>","Remarks");B:values("old"x4);C:reldates(0,0)
FIXTURE  wps_01.xlsx sheet=Plain cols=A:Label,B:Value rows=3 data=A:values("Other"x3);B:values("old"x3)
TEST     fixture=wps_01.xlsx call=SignWps(WPS) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(WPS![Value]3) = "R. Demo, Welding Engineer" AND VALUE(WPS![Value]4) = VALUE(WPS![Ref]4) AND VALUE(WPS![Value]2) = "old" AND VALUE(WPS![Value]5) = "old"
TEST     fixture=wps_01.xlsx call=SignWps(WPS) assert=COUNTA(WPS![Label]) = PRE_COUNTA(WPS![Label]) AND COUNTA(WPS![Ref]) = PRE_COUNTA(WPS![Ref]) AND VALUE(WPS![Label]3) = "Name >>" AND VALUE(WPS![Label]4) = "Date >>"
TEST     fixture=wps_01.xlsx call=SignWps(Plain) assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE(Plain![Value]2) = "old" AND VALUE(Plain![Value]3) = "old"
