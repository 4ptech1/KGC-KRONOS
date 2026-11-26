VCL 0.7
TASK
# ASSUME: the parent material block is rows 22 to 26 of the sheet WPS in both books
# ASSUME: the actual columns are C and E; the qualified column F and the # marks in D stay as they are
# ASSUME: the old WPS is read only; its values are copied as they are, blank cells copy as blank
GOAL     Copy the actual parent material values, rows 22 to 26 in columns C and E, from the old WPS into the new WPS form.
PROC     CopyParentMaterial
KIND     Function
MODULE   modWps
ARGS     wps As Worksheet, oldBook As Workbook
RET      Long
VAR      old As Worksheet
VAR      n As Long
PRE      SHEET_EXISTS($oldBook, "WPS"); SAME(CELL($wps, 22, 1), "Parent Material Type (Plate or Pipe)"); SAME(CELL($wps, 26, 1), "Outside Pipe Diameter (mm)")
MODIFIES $wps
OP       LET(old, SHEET($oldBook, "WPS"))
OP       LOOP_ROWS dir=down from=22 to=26 var=r
OP         SET_VALUE(CELL($wps, r, 3), CELL(old, r, 3))
OP         SET_VALUE(CELL($wps, r, 5), CELL(old, r, 5))
OP         LET(n, n + 1)
OP       RETURN(n)
ERR      standard
FIXTURE  wps_new_01.xlsx template=wps_form
FIXTURE  wps_old_01.xlsx template=wps_form_b
TEST     fixture=wps_new_01.xlsx read=wps_old_01.xlsx call=CopyParentMaterial(WPS, wps_old_01.xlsx) runs=2 assert=RESULT = 5 AND ERROR_COUNT = 0 AND VALUE(WPS!C22) = "Plate" AND VALUE(WPS!C23) = "SA 516 Gr 70" AND VALUE(WPS!E23) = "SA 516 Gr 70" AND VALUE(WPS!C25) = 12 AND VALUE(WPS!E26) = "NA"
TEST     fixture=wps_new_01.xlsx read=wps_old_01.xlsx call=CopyParentMaterial(WPS, wps_old_01.xlsx) assert=ERROR_COUNT = 0 AND VALUE(WPS!C6) = "DM-WPS-201" AND VALUE(WPS!F25) = "6 to 24 mm" AND COUNTA(WPS!A1:A50) = PRE_COUNTA(WPS!A1:A50)
