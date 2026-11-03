VCL 0.6
TASK
# ASSUME: NDE Report headers in row 1 include "Weld Seam No", "Report File" and "Status"
# ASSUME: the real sheet is "NDE Report"; the TESTs use NdeReport
# ASSUME: both folders exist; each report is copied directly into the Archive folder, never moved, and nothing is overwritten or deleted
# ASSUME: a row with a blank Weld Seam No is skipped; a blank or unsafe Report File gets Status "Bad name"
# ASSUME: Status holds the copy's result: Copied, Already there, Missing or Bad name, and also Conflict (a different file of that name is already in Archive), Locked or Too long
GOAL     Copy each NDE report file named in the NDE Report sheet from the Reports folder into the Archive folder and write what happened in the Status column.
PROC     ArchiveNdeReports
KIND     Function
MODULE   modNde
ARGS     nde As Worksheet, reports As String, archive As String
RET      Long
VAR      cSeam As Long
VAR      cFile As Long
VAR      cStatus As Long
VAR      status As String
VAR      copied As Long
PRE      COL($nde, "Weld Seam No") > 0; COL($nde, "Report File") > 0; COL($nde, "Status") > 0; FOLDER_EXISTS($reports); FOLDER_EXISTS($archive)
MODIFIES $nde["Status"]
MODIFIES $archive files
OP       LET(cSeam, COL($nde, "Weld Seam No"))
OP       LET(cFile, COL($nde, "Report File"))
OP       LET(cStatus, COL($nde, "Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($nde, cSeam) var=r
OP         SKIP_IF IS_EMPTY(CELL($nde, r, cSeam))
OP         COPY_FILE(PATH_JOIN($reports, SAFE_NAME(CELL($nde, r, cFile))), PATH_JOIN($archive, SAFE_NAME(CELL($nde, r, cFile))), into=status)
OP         SET_VALUE(CELL($nde, r, cStatus), status)
OP         IF status = "Copied" THEN LET(copied, copied + 1)
OP       RETURN(copied)
PERF     screen_off
ERR      standard
FIXTURE  nde_01.xlsx sheet=NdeReport cols=A:Weld Seam No,B:Report File,C:Status rows=10 data=A:values("W-101"x1,"W-102"x1,"W-103"x1,"W-104"x1,""x1,"W-105"x1,"W-106"x1,"W-107"x1,"W-108"x1,"W-109"x1)+boundary("W-101","W-102","W-103","W-104","","W-105","W-106","W-107","W-108","W-109");B:values("rpt_a.pdf"x2,"rpt_b.pdf"x1,"rpt_c.pdf"x1,"..\x.pdf"x1,"rpt_d.pdf"x1,"rpt_e.pdf"x1,""x1,"CON.pdf"x1,"rpt_a.pdf"x0)+boundary("rpt_a.pdf","rpt_b.pdf","rpt_c.pdf","..\x.pdf","rpt_a.pdf","rpt_d.pdf","rpt_e.pdf","","CON.pdf");C:values("old"x10)
FIXTURE  nde_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  nde_01.xlsx folder=Reports files=rpt_a.pdf:100,rpt_b.pdf:80,rpt_d.pdf:60,rpt_e.pdf:50 locked=rpt_e.pdf
FIXTURE  nde_01.xlsx folder=Archive files=rpt_b.pdf:80,rpt_d.pdf:61
TEST     fixture=nde_01.xlsx call=ArchiveNdeReports(NdeReport, Reports, Archive) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0
TEST     fixture=nde_01.xlsx call=ArchiveNdeReports(NdeReport, Reports, Archive) assert=RESULT = 1
TEST     fixture=nde_01.xlsx call=ArchiveNdeReports(Notes, Reports, Archive) assert=RESULT = -1 AND ERROR_COUNT = 1 AND FILE_COUNT("Archive") = PRE_FILE_COUNT("Archive")
