VCL 0.7
TASK
# ASSUME: the sheet is "WPS" in both workbooks, laid out as the practice form; its labels are in column A, so they are checked by their cells in PRE
# ASSUME: the parent material block is rows 22 to 26 (type, designation, group, thickness, outside pipe diameter); the Actual cells are columns C and E
# ASSUME: the "#" column D and the Qualified column F are not copied; the old WPS is only read
# ASSUME: a blank cell in the old WPS is not copied, so the new form keeps what it had there
# ASSUME: the TESTs use plain sheet FIXTURES for the old WPS (invented values in the form's cells), and the template for the new form
GOAL     Copy the parent material block, the actual columns for both members, from the old WPS into the new WPS form and only read the old one.
PROC     CopyParentMaterial
KIND     Function
MODULE   modWps
ARGS     wps As Worksheet, oldBook As Workbook
RET      Long
VAR      src As Worksheet
VAR      n As Long
PRE      CELL($wps, 22, 1) = "Parent Material Type (Plate or Pipe)"; CELL($wps, 26, 1) = "Outside Pipe Diameter (mm)"; CELL(SHEET($oldBook, "WPS"), 22, 1) = "Parent Material Type (Plate or Pipe)"; CELL(SHEET($oldBook, "WPS"), 26, 1) = "Outside Pipe Diameter (mm)"
MODIFIES $wps
OP       LET(src, SHEET($oldBook, "WPS"))
OP       LOOP_ROWS dir=down from=22 to=26 var=r
OP         IF NOT IS_EMPTY(CELL(src, r, 3))
OP           SET_VALUE(CELL($wps, r, 3), CELL(src, r, 3))
OP           LET(n, n + 1)
OP         IF NOT IS_EMPTY(CELL(src, r, 5))
OP           SET_VALUE(CELL($wps, r, 5), CELL(src, r, 5))
OP           LET(n, n + 1)
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  new_01.xlsx template=wps_form
FIXTURE  old_01.xlsx sheet=WPS cols=A:Label,B:Mark,C:Actual,D:Hash,E:Actual2 rows=30 data=A:text(12)+boundary("","","","","","","","","","","","","","","","","","","","","Parent Material Type (Plate or Pipe)","Parent Material Designation","Parent Material Group (ISO/TR 15608)","Material Thickness (mm)","Outside Pipe Diameter (mm)");B:text(4);C:text(12)+boundary("","","","","","","","","","","","","","","","","","","","","Pipe","SA 106 Gr B","1.1","8","508");D:text(1);E:text(12)+boundary("","","","","","","","","","","","","","","","","","","","","Pipe","SA 335 P11","5.1","9","610")
FIXTURE  old_02.xlsx sheet=WPS cols=A:Label,B:Mark,C:Actual,D:Hash,E:Actual2 rows=30 data=A:text(12)+boundary("","","","","","","","","","","","","","","","","","","","","Parent Material Type (Plate or Pipe)","Parent Material Designation","Parent Material Group (ISO/TR 15608)","Material Thickness (mm)","Outside Pipe Diameter (mm)");B:text(4);C:text(12)+boundary("","","","","","","","","","","","","","","","","","","","","Pipe","SA 106 Gr B","","8","508");D:text(1);E:text(12)+boundary("","","","","","","","","","","","","","","","","","","","","Pipe","","5.1","9","610")
TEST     fixture=new_01.xlsx read=old_01.xlsx call=CopyParentMaterial(WPS, old_01.xlsx) runs=2 assert=RESULT = 10 AND ERROR_COUNT = 0 AND VALUE(WPS!C22) = "Pipe" AND VALUE(WPS!C23) = "SA 106 Gr B" AND VALUE(WPS!C24) = "1.1" AND VALUE(WPS!C25) = "8" AND VALUE(WPS!C26) = "508" AND VALUE(WPS!E22) = "Pipe" AND VALUE(WPS!E23) = "SA 335 P11" AND VALUE(WPS!E24) = "5.1" AND VALUE(WPS!E25) = "9" AND VALUE(WPS!E26) = "610"
TEST     fixture=new_01.xlsx read=old_02.xlsx call=CopyParentMaterial(WPS, old_02.xlsx) assert=RESULT = 8 AND ERROR_COUNT = 0 AND VALUE(WPS!C24) = PRE_VALUE(WPS!C24) AND VALUE(WPS!E23) = PRE_VALUE(WPS!E23) AND VALUE(WPS!C23) = "SA 106 Gr B" AND VALUE(WPS!E24) = "5.1"
TEST     fixture=new_01.xlsx read=old_01.xlsx call=CopyParentMaterial(WPS, old_01.xlsx) assert=VALUE(WPS!D22) = PRE_VALUE(WPS!D22) AND VALUE(WPS!F22) = PRE_VALUE(WPS!F22) AND VALUE(WPS!F25) = PRE_VALUE(WPS!F25) AND VALUE(WPS!C21) = PRE_VALUE(WPS!C21) AND VALUE(WPS!C6) = PRE_VALUE(WPS!C6) AND VALUE(WPS!C27) = PRE_VALUE(WPS!C27) AND VALUE(WPS!A22) = PRE_VALUE(WPS!A22)
