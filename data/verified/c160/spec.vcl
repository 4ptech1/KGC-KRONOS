VCL 0.7
TASK
# ASSUME: The practice form is the sheet WPS; the date is written as "DATE:- 25/11/2026" and the revision as "WPS Revision : R1", in the cells that hold those labels
# ASSUME: Page 1 labels are H2 and H3, page 2 labels are H35 and H36; nothing else is touched
GOAL     Change the date to 25/11/2026 and the revision to R1 in the header of page 1 and again in the header of page 2 of the WPS form.
PROC     SetWpsDateAndRevision
KIND     Function
MODULE   modWps
ARGS     ws As Worksheet
RET      Long
VAR      n As Long
PRE      LEFT(CELL($ws, 2, 8), 5) = "DATE:"; LEFT(CELL($ws, 3, 8), 7) = "WPS Rev"; LEFT(CELL($ws, 35, 8), 5) = "DATE:"; LEFT(CELL($ws, 36, 8), 7) = "WPS Rev"
MODIFIES $ws
OP       SET_VALUE(CELL($ws, 2, 8), "DATE:- 25/11/2026")
OP       SET_VALUE(CELL($ws, 3, 8), "WPS Revision : R1")
OP       SET_VALUE(CELL($ws, 35, 8), "DATE:- 25/11/2026")
OP       SET_VALUE(CELL($ws, 36, 8), "WPS Revision : R1")
OP       LET(n, 4)
OP       RETURN(n)
ERR      standard
FIXTURE  wps_01.xlsx template=wps_form
TEST     fixture=wps_01.xlsx call=SetWpsDateAndRevision(WPS) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND VALUE(WPS!H2) = "DATE:- 25/11/2026" AND VALUE(WPS!H3) = "WPS Revision : R1" AND VALUE(WPS!H35) = "DATE:- 25/11/2026" AND VALUE(WPS!H36) = "WPS Revision : R1"
TEST     fixture=wps_01.xlsx call=SetWpsDateAndRevision(WPS) assert=VALUE(WPS!H1) = "FORMAT NO:- QF/DM/31" AND VALUE(WPS!C6) = "DM-WPS-201" AND VALUE(WPS!B50) = "02/11/2026" AND COUNTA(WPS!A1:AC50) = PRE_COUNTA(WPS!A1:AC50) AND VALUE(WPS!H34) = "FORMAT NO:- QF/DM/31"
