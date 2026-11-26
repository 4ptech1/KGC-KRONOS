VCL 0.7
TASK
# ASSUME: The parent material block is rows 22 to 26 of the WPS sheet; its actual columns are C and E (column D holds the # marker and is left alone)
# ASSUME: The old WPS has a sheet named WPS with the same layout, and is only read
# ASSUME: A cell counts as copied when the old cell is not blank; a blank old cell leaves the new cell as it is
GOAL     Copy the parent material block's actual columns from the old WPS workbook into the new WPS form and return how many cells were copied.
PROC     CopyParentMaterial
KIND     Function
MODULE   modWps
ARGS     wps As Worksheet, oldBook As Workbook
RET      Long
VAR      old As Worksheet
VAR      copied As Long
PRE      SHEET_EXISTS($oldBook, "WPS"); SAME(CELL($wps, 22, 1), "Parent Material Type (Plate or Pipe)"); SAME(CELL($wps, 26, 1), "Outside Pipe Diameter (mm)")
MODIFIES $wps
OP       LET(old, SHEET($oldBook, "WPS"))
OP       LOOP_ROWS dir=down from=22 to=26 var=r
OP         LOOP_COLS dir=right from=3 to=5 var=c
OP           SKIP_IF c = 4
OP           SKIP_IF IS_EMPTY(CELL(old, r, c))
OP           SET_VALUE(CELL($wps, r, c), CELL(old, r, c))
OP           LET(copied, copied + 1)
OP       RETURN(copied)
ERR      standard
FIXTURE  wps_new.xlsx template=wps_form
FIXTURE  wps_old.xlsx template=wps_form_b
TEST     fixture=wps_new.xlsx read=wps_old.xlsx call=CopyParentMaterial(WPS, wps_old.xlsx) runs=2 assert=RESULT = 10 AND ERROR_COUNT = 0 AND VALUE(WPS!C23) = "SA 516 Gr 70" AND VALUE(WPS!E25) = 12
TEST     fixture=wps_new.xlsx read=wps_old.xlsx call=CopyParentMaterial(WPS, wps_old.xlsx) assert=ERROR_COUNT = 0 AND COUNTA(WPS!A22:A26) = PRE_COUNTA(WPS!A22:A26) AND COUNTA(WPS!F22:F26) = PRE_COUNTA(WPS!F22:F26) AND VALUE(WPS!C22) = "Plate"
