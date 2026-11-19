VCL 0.6
TASK
# ASSUME: Welder Quals has headers in row 1 including "Welder ID" and "Folder Status"; the Welder Certs folder already exists and is passed in as a path
# ASSUME: a folder is named after the Welder ID with stray spaces trimmed and inner spaces made one; a welder on several rows gets one folder, and the later rows say Already there
# ASSUME: Folder Status says Made when the folder was created, Already there when it existed, and Bad name for a blank or unsafe Welder ID
# ASSUME: any other outcome (a file with that name in the way, a path too long) is written as the tool reports it, so no row is left unexplained
# ASSUME: nothing is deleted, overwritten or moved; only the Folder Status column of Welder Quals changes
GOAL     Make a folder named after each Welder ID inside the Welder Certs folder and write Made, Already there or Bad name in Folder Status on every row.
PROC     MakeWelderCertFolders
KIND     Function
MODULE   modWelderCerts
ARGS     quals As Worksheet, certs As String
RET      Long
VAR      cId As Long
VAR      cStatus As Long
VAR      status As String
VAR      made As Long
PRE      COL($quals, "Welder ID") > 0; COL($quals, "Folder Status") > 0; FOLDER_EXISTS($certs)
MODIFIES $quals["Folder Status"]
MODIFIES $certs files
OP       LET(cId, COL($quals, "Welder ID"))
OP       LET(cStatus, COL($quals, "Folder Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($quals, cId) var=r
OP         MAKE_FOLDER(PATH_JOIN($certs, SAFE_NAME(CELL($quals, r, cId))), into=status)
OP         SELECT_CASE status
OP           CASE "Created"
OP             SET_VALUE(CELL($quals, r, cStatus), "Made")
OP             LET(made, made + 1)
OP           CASE "Exists"
OP             SET_VALUE(CELL($quals, r, cStatus), "Already there")
OP           CASE_ELSE
OP             SET_VALUE(CELL($quals, r, cStatus), status)
OP       RETURN(made)
PERF     screen_off
ERR      standard
FIXTURE  certs_01.xlsx sheet=Quals cols=A:Welder ID,B:Process,C:Folder Status rows=10 data=A:values("W-01"x2,"W-02"x1,""x1,"..\x"x1,"W-03 "x1,"a/b"x1,"CON"x1,"W-09"x1,"W-10"x1)+boundary("W-01","W-02","W-01","","..\x","W-03 ","a/b","CON","W-09","W-10");B:values("GTAW"x10);C:values("old"x10)
FIXTURE  certs_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  certs_01.xlsx folder=Certs files=W-02\scan.txt:20,W-09:5,W-11\scan.txt:30
TEST     fixture=certs_01.xlsx call=MakeWelderCertFolders(Quals, Certs) runs=2 assert=ERROR_COUNT = 0 AND RESULT = 0 AND COUNTIF(Quals![Folder Status], "Already there") = 5 AND COUNTIF(Quals![Folder Status], "Bad name") = 4 AND COUNTIF(Quals![Folder Status], "Blocked by file") = 1 AND COUNTIF(Quals![Folder Status], "old") = 0 AND FILE_COUNT("Certs") = PRE_FILE_COUNT("Certs")
TEST     fixture=certs_01.xlsx call=MakeWelderCertFolders(Quals, Certs) assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Quals![Folder Status], "Made") = 3 AND COUNTIF(Quals![Folder Status], "Already there") = 2 AND VALUE(Quals![Folder Status]2) = "Made" AND VALUE(Quals![Folder Status]3) = "Already there" AND VALUE(Quals![Folder Status]4) = "Already there" AND VALUE(Quals![Folder Status]5) = "Bad name" AND VALUE(Quals![Folder Status]7) = "Made" AND VALUE(Quals![Folder Status]10) = "Blocked by file" AND FILE_COUNT("Certs") = PRE_FILE_COUNT("Certs") AND COUNTA(Quals![Welder ID]) = PRE_COUNTA(Quals![Welder ID]) AND COUNTA(Quals![Process]) = PRE_COUNTA(Quals![Process])
TEST     fixture=certs_01.xlsx call=MakeWelderCertFolders(Notes, Certs) assert=RESULT = -1 AND ERROR_COUNT = 1 AND FILE_COUNT("Certs") = PRE_FILE_COUNT("Certs")
