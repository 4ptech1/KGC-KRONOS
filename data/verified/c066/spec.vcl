VCL 0.6
TASK
# ASSUME: Files is the sheet that is rebuilt on every run; it already exists and is passed in, and everything on it is cleared first
# ASSUME: jobFolder is the full path of a folder that exists; only the files directly in it are listed, not those in folders below it
# ASSUME: names are listed in name order, ignoring case; hidden and system files are left out
# ASSUME: the count is returned and also shown in a message
GOAL     List the names of the files in the job folder on the Files sheet under the heading File Name, rebuilding the sheet on every run, and tell how many files there are.
PROC     ListJobFiles
KIND     Function
MODULE   modJobFiles
ARGS     filesSheet As Worksheet, jobFolder As String
RET      Long
VAR      names As Collection
VAR      n As Long
VAR      outRow As Long
PRE      FOLDER_EXISTS($jobFolder)
MODIFIES $filesSheet
OP       LIST_FILES($jobFolder, "*", into=names)
OP       CLEAR(USED_RANGE($filesSheet), what=all)
OP       WRITE_ROW($filesSheet, 1, "File Name")
OP       LET(outRow, 2)
OP       FOR_EACH var=nm in=ITEMS(names)
OP         WRITE_ROW($filesSheet, outRow, nm)
OP         LET(outRow, outRow + 1)
OP       LET(n, COUNT(names))
OP       MSG("Files in the job folder: " & n)
OP       RETURN(n)
ERR      standard
FIXTURE  jobs_01.xlsx sheet=Files cols=A:File Name rows=0 stale=6
FIXTURE  jobs_01.xlsx folder=JobFolder files=Zed.txt:10,alpha.pdf:20,Bravo.xlsx:30,notes_2.txt:5,sub\inner.txt:7 locked=Zed.txt
FIXTURE  jobs_02.xlsx sheet=Files cols=A:File Name rows=0 stale=3
FIXTURE  jobs_02.xlsx folder=JobFolder files=sub\inner.txt:7
TEST     fixture=jobs_01.xlsx call=ListJobFiles(Files, JobFolder) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND ROWS_USED(Files) = 5 AND COUNTA(Files![File Name]) = 5 AND VALUE(Files![File Name]1) = "File Name" AND VALUE(Files![File Name]2) = "alpha.pdf" AND VALUE(Files![File Name]3) = "Bravo.xlsx" AND VALUE(Files![File Name]4) = "notes_2.txt" AND VALUE(Files![File Name]5) = "Zed.txt" AND COUNTIF(Files![File Name], "stale") = 0 AND FILL_COUNT(Files![File Name], #FF0000) = 0
TEST     fixture=jobs_01.xlsx call=ListJobFiles(Files, JobFolder) assert=RESULT = 4 AND ERROR_COUNT = 0 AND FILE_COUNT("JobFolder") = PRE_FILE_COUNT("JobFolder") AND ROWS_USED(Files) = RESULT + 1
TEST     fixture=jobs_02.xlsx call=ListJobFiles(Files, JobFolder) assert=RESULT = 0 AND ERROR_COUNT = 0 AND ROWS_USED(Files) = 1 AND COUNTA(Files![File Name]) = 1 AND VALUE(Files![File Name]1) = "File Name" AND COUNTIF(Files![File Name], "stale") = 0 AND FILE_COUNT("JobFolder") = PRE_FILE_COUNT("JobFolder")
TEST     fixture=jobs_01.xlsx call=ListJobFiles(Files, "NoSuchFolder") assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(Files![File Name]) = PRE_COUNTA(Files![File Name]) AND COUNTIF(Files![File Name], "stale") = PRE_COUNTIF(Files![File Name], "stale")
