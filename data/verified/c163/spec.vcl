VCL 0.7
TASK
# ASSUME: the sheet is "WPS" in both workbooks, laid out as the practice form; its labels are in column A, so they are checked by their cells in PRE
# ASSUME: copied from the old WPS: company name C4, address C5, welding process C8, degree of mechanization C9, filler designation actual C39 and qualified F39
# ASSUME: the new WPS number DM-WPS-203 is written into C6; the revision, PQR number and the rest of the form are not touched
# ASSUME: a blank cell in the old WPS is not copied, so the new form keeps what it had there; the old WPS is only read
# ASSUME: the TESTs use plain sheet FIXTURES for the old WPS (invented values in the form's cells), and the template for the new form
GOAL     Make a new WPS from the old one by copying company, process, mechanization and filler designation and writing the new number DM-WPS-203.
PROC     MakeNewWps
KIND     Function
MODULE   modWps
ARGS     wps As Worksheet, oldBook As Workbook
RET      Long
VAR      src As Worksheet
VAR      n As Long
PRE      CELL($wps, 4, 1) = "Company Name :"; CELL($wps, 6, 1) = "Welding Procedure Specification No. :"; CELL($wps, 39, 1) = "Filler Material Designation :"; CELL(SHEET($oldBook, "WPS"), 4, 1) = "Company Name :"; CELL(SHEET($oldBook, "WPS"), 39, 1) = "Filler Material Designation :"
MODIFIES $wps
OP       LET(src, SHEET($oldBook, "WPS"))
OP       LOOP_ROWS dir=down from=4 to=9 var=r
OP         SKIP_IF r = 6 OR r = 7
OP         SKIP_IF IS_EMPTY(CELL(src, r, 3))
OP         SET_VALUE(CELL($wps, r, 3), CELL(src, r, 3))
OP         LET(n, n + 1)
OP       IF NOT IS_EMPTY(CELL(src, 39, 3))
OP         SET_VALUE(CELL($wps, 39, 3), CELL(src, 39, 3))
OP         LET(n, n + 1)
OP       IF NOT IS_EMPTY(CELL(src, 39, 6))
OP         SET_VALUE(CELL($wps, 39, 6), CELL(src, 39, 6))
OP         LET(n, n + 1)
OP       SET_VALUE(CELL($wps, 6, 3), "DM-WPS-203")
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  new_01.xlsx template=wps_form
FIXTURE  old_01.xlsx sheet=WPS cols=A:Label,B:Mark,C:Actual,D:Hash,E:Gap,F:Qualified rows=45 data=A:text(12)+boundary("","","Company Name :","Company Address :","Welding Procedure Specification No. :","","Welding Process :","Degree of mechanization :","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","Filler Material Designation :","","","");B:text(4);C:text(12)+boundary("","Decoy 3","Alpha Works Pvt Ltd","Plot 7, Sample Park","OLD-WPS-9","OLD-PQR-9","SMAW (111)","Automatic","Decoy 10","","","","","","","","","","","","","","","","","","","","","","","","","","","","Decoy 38","AWS A5.1 E7018","Decoy 40","Decoy 41","");D:text(1);E:text(3);F:text(12)+boundary("","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","AWS A5.1 E7018, E7016","Decoy F40","","")
FIXTURE  old_02.xlsx sheet=WPS cols=A:Label,B:Mark,C:Actual,D:Hash,E:Gap,F:Qualified rows=45 data=A:text(12)+boundary("","","Company Name :","Company Address :","Welding Procedure Specification No. :","","Welding Process :","Degree of mechanization :","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","Filler Material Designation :","","","");B:text(4);C:text(12)+boundary("","Decoy 3","Alpha Works Pvt Ltd","","OLD-WPS-9","OLD-PQR-9","SMAW (111)","Automatic","Decoy 10","","","","","","","","","","","","","","","","","","","","","","","","","","","","Decoy 38","","Decoy 40","Decoy 41","");D:text(1);E:text(3);F:text(12)+boundary("","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","","AWS A5.1 E7018, E7016","Decoy F40","","")
TEST     fixture=new_01.xlsx read=old_01.xlsx call=MakeNewWps(WPS, old_01.xlsx) runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND VALUE(WPS!C4) = "Alpha Works Pvt Ltd" AND VALUE(WPS!C5) = "Plot 7, Sample Park" AND VALUE(WPS!C8) = "SMAW (111)" AND VALUE(WPS!C9) = "Automatic" AND VALUE(WPS!C39) = "AWS A5.1 E7018" AND VALUE(WPS!F39) = "AWS A5.1 E7018, E7016" AND VALUE(WPS!C6) = "DM-WPS-203"
TEST     fixture=new_01.xlsx read=old_02.xlsx call=MakeNewWps(WPS, old_02.xlsx) assert=RESULT = 4 AND ERROR_COUNT = 0 AND VALUE(WPS!C5) = PRE_VALUE(WPS!C5) AND VALUE(WPS!C39) = PRE_VALUE(WPS!C39) AND VALUE(WPS!C4) = "Alpha Works Pvt Ltd" AND VALUE(WPS!F39) = "AWS A5.1 E7018, E7016" AND VALUE(WPS!C6) = "DM-WPS-203"
TEST     fixture=new_01.xlsx read=old_01.xlsx call=MakeNewWps(WPS, old_01.xlsx) assert=VALUE(WPS!C7) = PRE_VALUE(WPS!C7) AND VALUE(WPS!C10) = PRE_VALUE(WPS!C10) AND VALUE(WPS!F6) = PRE_VALUE(WPS!F6) AND VALUE(WPS!F7) = PRE_VALUE(WPS!F7) AND VALUE(WPS!C12) = PRE_VALUE(WPS!C12) AND VALUE(WPS!C13) = PRE_VALUE(WPS!C13) AND VALUE(WPS!C37) = PRE_VALUE(WPS!C37) AND VALUE(WPS!C40) = PRE_VALUE(WPS!C40) AND VALUE(WPS!A4) = PRE_VALUE(WPS!A4) AND VALUE(WPS!C1) = PRE_VALUE(WPS!C1) AND VALUE(WPS!H3) = PRE_VALUE(WPS!H3)
