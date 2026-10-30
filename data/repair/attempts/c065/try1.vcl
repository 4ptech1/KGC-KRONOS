VCL 0.6
TASK
# ASSUME: Log is the sheet that gets the answer in cell A1; whatever is in A1 is replaced
# ASSUME: source and archive are full paths of folders that already exist; the file is weld_map.pdf in both
# ASSUME: a file named weld_map.pdf of any size already in archive (any letter case) counts as Exists and is never overwritten
# ASSUME: Exists is checked first, so it is written even when weld_map.pdf is also missing from source
# ASSUME: any other result, such as Locked, is written to A1 as it is
GOAL     Copy weld_map.pdf from the Source folder to the Archive folder unless it is already there, and write Copied, Exists or Missing in cell A1 of the Log sheet.
PROC     CopyWeldMapToArchive
KIND     Function
MODULE   modArchive
ARGS     logSheet As Worksheet, source As String, archive As String
RET      Long
VAR      found As Collection
VAR      status As String
VAR      copied As Long
PRE      FOLDER_EXISTS($source); FOLDER_EXISTS($archive)
MODIFIES $logSheet
MODIFIES $archive files
OP       LIST_FILES($archive, "weld_map.pdf", into=found)
OP       IF COUNT(found) > 0
OP         SET_VALUE(CELL($logSheet, 1, 1), "Exists")
OP       ELSE
OP         COPY_FILE(PATH_JOIN($source, "weld_map.pdf"), PATH_JOIN($archive, "weld_map.pdf"), into=status)
OP         SET_VALUE(CELL($logSheet, 1, 1), status)
OP         IF status = "Copied" THEN LET(copied, 1)
OP       RETURN(copied)
ERR      standard
FIXTURE  weld_01.xlsx sheet=Log cols=A:Result rows=0
FIXTURE  weld_01.xlsx folder=Source files=weld_map.pdf:120,rev_a.txt:10
FIXTURE  weld_01.xlsx folder=Archive files=history.txt:5,weld_map.pdf.bak:7
FIXTURE  weld_02.xlsx sheet=Log cols=A:Result rows=0
FIXTURE  weld_02.xlsx folder=Source files=weld_map.pdf:"new copy",rev_a.txt:10
FIXTURE  weld_02.xlsx folder=Archive files=WELD_MAP.PDF:"old copy",history.txt:5
FIXTURE  weld_03.xlsx sheet=Log cols=A:Result rows=0
FIXTURE  weld_03.xlsx folder=Source files=weld_map.dwg:120,rev_a.txt:10
FIXTURE  weld_03.xlsx folder=Archive files=history.txt:5
FIXTURE  weld_04.xlsx sheet=Log cols=A:Result rows=0
FIXTURE  weld_04.xlsx folder=Source files=weld_map.pdf:120,rev_a.txt:10 locked=weld_map.pdf
FIXTURE  weld_04.xlsx folder=Archive files=history.txt:5
TEST     fixture=weld_01.xlsx call=CopyWeldMapToArchive(Log, Source, Archive) assert=RESULT = 1 AND ERROR_COUNT = 0 AND VALUE(Log!A1) = "Copied" AND FILE_EXISTS("Archive\weld_map.pdf") AND FILE_COUNT("Archive") = PRE_FILE_COUNT("Archive") + 1 AND FILE_COUNT("Source") = PRE_FILE_COUNT("Source")
TEST     fixture=weld_01.xlsx call=CopyWeldMapToArchive(Log, Source, Archive) runs=2 assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE(Log!A1) = "Exists" AND FILE_COUNT("Archive") = PRE_FILE_COUNT("Archive") + 1
TEST     fixture=weld_02.xlsx call=CopyWeldMapToArchive(Log, Source, Archive) assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE(Log!A1) = "Exists" AND FILE_TEXT("Archive\WELD_MAP.PDF") = "old copy" AND FILE_COUNT("Archive") = PRE_FILE_COUNT("Archive") AND FILE_COUNT("Source") = PRE_FILE_COUNT("Source")
TEST     fixture=weld_03.xlsx call=CopyWeldMapToArchive(Log, Source, Archive) assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE(Log!A1) = "Missing" AND FILE_COUNT("Archive") = PRE_FILE_COUNT("Archive")
TEST     fixture=weld_04.xlsx call=CopyWeldMapToArchive(Log, Source, Archive) assert=RESULT = 0 AND ERROR_COUNT = 0 AND VALUE(Log!A1) = "Locked" AND FILE_EXISTS("Archive\weld_map.pdf") = 0 AND FILE_COUNT("Archive") = PRE_FILE_COUNT("Archive")
TEST     fixture=weld_01.xlsx call=CopyWeldMapToArchive(Log, Source, "NoSuchFolder") assert=RESULT = -1 AND ERROR_COUNT = 1 AND FILE_COUNT("Archive") = PRE_FILE_COUNT("Archive")
