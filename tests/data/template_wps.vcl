VCL 0.7
TASK
# ASSUME: The WPS form has its revision line in H3 and H36 and its date line in H2 and H35, one on each page.
GOAL     Set the revision and the date on both pages of the practice WPS form.
PROC     SetWpsRevision
KIND     Function
MODULE   modWps
ARGS     ws As Worksheet, rev As String, dateText As String
RET      Long
PRE      SAME(CELL($ws, 6, 1), "Welding Procedure Specification No. :")
MODIFIES $ws
OP       SET_VALUE(RANGE($ws, "H2"), "DATE:- " & $dateText)
OP       SET_VALUE(RANGE($ws, "H35"), "DATE:- " & $dateText)
OP       SET_VALUE(RANGE($ws, "H3"), "WPS Revision : " & $rev)
OP       SET_VALUE(RANGE($ws, "H36"), "WPS Revision : " & $rev)
OP       RETURN(4)
ERR      standard
FIXTURE  wps_01.xlsx template=wps_form
TEST     fixture=wps_01.xlsx call=SetWpsRevision(WPS, "R1", "20/11/2026") runs=2 assert=RESULT = 4 AND VALUE(WPS!H3) = "WPS Revision : R1" AND VALUE(WPS!H36) = "WPS Revision : R1" AND VALUE(WPS!H35) = "DATE:- 20/11/2026" AND ERROR_COUNT = 0
TEST     fixture=wps_01.xlsx call=SetWpsRevision(WPS, "R2", "21/11/2026") assert=COUNTA(WPS!A1:I50) = PRE_COUNTA(WPS!A1:I50) AND FORMULA(WPS!R29) = "=$N$25*N29*L29*60/(Q29*1000)" AND VALUE(WPS!C6) = "DM-WPS-201"
