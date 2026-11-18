VCL 0.6
TASK
# ASSUME: Job Tracker headers in row 1 are Job ID, Status and Folder (the TEST sheet is called Tracker)
# ASSUME: the Jobs folder path is the String argument jobs and it already exists; nothing is deleted, moved or overwritten
# ASSUME: a job is Closed when its Status is Closed, ignoring case and extra spaces; its Folder cell is emptied
# ASSUME: Folder says Made when the job folder or its Drawings folder was created, Already there when both existed, Bad name when the Job ID is not a safe folder name
# ASSUME: any other outcome (for example a file with the job's name) is written as MAKE_FOLDER's own status text; a blank Job ID row is skipped and left as it is
GOAL     Make a folder named after each open job with a Drawings folder inside it, and write Made, Already there or Bad name in the Folder column.
PROC     MakeJobFolders
KIND     Function
MODULE   modJobFolders
ARGS     tracker As Worksheet, jobs As String
RET      Long
VAR      cJob As Long
VAR      cStatus As Long
VAR      cFolder As Long
VAR      st As String
VAR      st2 As String
VAR      made As Long
PRE      COL($tracker, "Job ID") > 0; COL($tracker, "Status") > 0; COL($tracker, "Folder") > 0; FOLDER_EXISTS($jobs)
MODIFIES $tracker["Folder"]
MODIFIES $jobs files
OP       LET(cJob, COL($tracker, "Job ID"))
OP       LET(cStatus, COL($tracker, "Status"))
OP       LET(cFolder, COL($tracker, "Folder"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($tracker, cJob) var=r
OP         SKIP_IF IS_EMPTY(CELL($tracker, r, cJob))
OP         IF SAME(CELL($tracker, r, cStatus), "Closed")
OP           CLEAR(CELL($tracker, r, cFolder), what=contents)
OP         ELSE
OP           MAKE_FOLDER(PATH_JOIN($jobs, SAFE_NAME(CELL($tracker, r, cJob))), into=st)
OP           IF st = "Created" OR st = "Exists"
OP             MAKE_FOLDER(PATH_JOIN($jobs, SAFE_NAME(CELL($tracker, r, cJob)), "Drawings"), into=st2)
OP             IF st2 = "Created" OR st2 = "Exists"
OP               IF st = "Created" OR st2 = "Created"
OP                 SET_VALUE(CELL($tracker, r, cFolder), "Made")
OP                 LET(made, made + 1)
OP               ELSE
OP                 SET_VALUE(CELL($tracker, r, cFolder), "Already there")
OP             ELSE
OP               SET_VALUE(CELL($tracker, r, cFolder), st2)
OP           ELSE
OP             SET_VALUE(CELL($tracker, r, cFolder), st)
OP       RETURN(made)
PERF     screen_off
ERR      standard
FIXTURE  folders_01.xlsx sheet=Tracker cols=A:Job ID,B:Status,C:Folder rows=8 data=A:values("J-1"x1,"J-2"x1,"J-3"x1,"J-4"x1,"A*B"x1,""x1,"J-5"x1,"J-6"x1)+boundary("J-1","J-2","J-3","J-4","A*B","","J-5","J-6");B:values("Open"x5,"In Progress"x1,"Closed"x1," closed "x1)+boundary("Open","In Progress","Closed","Open","Open","Open","Open"," closed ");C:values(""x5,"old"x3)+boundary("","","old","","","old","","old")
FIXTURE  folders_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  folders_01.xlsx folder=Jobs files=J-2\Drawings\x.txt:10,J-4\n.txt:5,J-5:10
TEST     fixture=folders_01.xlsx call=MakeJobFolders(Tracker, Jobs) assert=RESULT = 2 AND ERROR_COUNT = 0 AND COUNTIF(Tracker![Folder], "Made") = 2 AND COUNTIF(Tracker![Folder], "Already there") = 1 AND COUNTIF(Tracker![Folder], "Bad name") = 1 AND COUNTIF(Tracker![Folder], "Blocked by file") = 1 AND COUNTIF(Tracker![Folder], "old") = 1 AND COUNTA(Tracker![Folder]) = 7 AND VALUE(Tracker![Folder]2) = "Made" AND VALUE(Tracker![Folder]3) = "Already there" AND VALUE(Tracker![Folder]5) = "Made" AND FILE_COUNT("Jobs") = PRE_FILE_COUNT("Jobs")
TEST     fixture=folders_01.xlsx call=MakeJobFolders(Tracker, Jobs) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(Tracker![Folder], "Already there") = 3 AND COUNTIF(Tracker![Folder], "Made") = 0 AND COUNTIF(Tracker![Folder], "Bad name") = 1 AND COUNTA(Tracker![Job ID]) = PRE_COUNTA(Tracker![Job ID]) AND FILE_COUNT("Jobs") = PRE_FILE_COUNT("Jobs")
TEST     fixture=folders_01.xlsx call=MakeJobFolders(Notes, Jobs) assert=RESULT = -1 AND ERROR_COUNT = 1 AND FILE_COUNT("Jobs") = PRE_FILE_COUNT("Jobs")
