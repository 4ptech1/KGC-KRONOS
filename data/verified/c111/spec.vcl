VCL 0.6
TASK
# ASSUME: GRN headers in row 1 include "GRN No", "Delivery Note File" and "Copy Status"; the real sheet is "GRN" and the TESTs use Grn
# ASSUME: inbox and archive are full paths of two folders that already exist; the macro stops without changes if either is missing
# ASSUME: each file is copied, never moved, deleted or overwritten; the archive folder for a GRN is named after its GRN No and made when missing
# ASSUME: a row with a blank GRN No is skipped and its Copy Status is left as it is
# ASSUME: a GRN No or Delivery Note File that is not safe as one folder or file name (or a blank Delivery Note File) gets Copy Status "Bad name"
# ASSUME: any other result of the copy (Conflict, Locked, Too long, ...) or of making the folder is written to Copy Status as it is
# ASSUME: the macro returns how many rows are now in the archive (Copied or Already there)
GOAL     Copy each GRN's delivery note from the inbox folder into a folder named after the GRN No inside the archive folder and write what happened in Copy Status.
PROC     ArchiveDeliveryNotes
KIND     Function
MODULE   modGrn
ARGS     grn As Worksheet, inbox As String, archive As String
RET      Long
VAR      cGrn As Long
VAR      cFile As Long
VAR      cStatus As Long
VAR      status As String
VAR      archived As Long
PRE      COL($grn, "GRN No") > 0; COL($grn, "Delivery Note File") > 0; COL($grn, "Copy Status") > 0; FOLDER_EXISTS($inbox); FOLDER_EXISTS($archive)
MODIFIES $grn["Copy Status"]
MODIFIES $archive files
OP       LET(cGrn, COL($grn, "GRN No"))
OP       LET(cFile, COL($grn, "Delivery Note File"))
OP       LET(cStatus, COL($grn, "Copy Status"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($grn, cGrn) var=r
OP         SKIP_IF IS_EMPTY(CELL($grn, r, cGrn))
OP         IF SAFE_NAME(CELL($grn, r, cGrn)) = "" OR SAFE_NAME(CELL($grn, r, cFile)) = ""
OP           SET_VALUE(CELL($grn, r, cStatus), "Bad name")
OP         ELSE
OP           MAKE_FOLDER(PATH_JOIN($archive, SAFE_NAME(CELL($grn, r, cGrn))), into=status)
OP           IF status = "Created" OR status = "Exists"
OP             COPY_FILE(PATH_JOIN($inbox, SAFE_NAME(CELL($grn, r, cFile))), PATH_JOIN($archive, SAFE_NAME(CELL($grn, r, cGrn)), SAFE_NAME(CELL($grn, r, cFile))), into=status)
OP           SET_VALUE(CELL($grn, r, cStatus), status)
OP           IF status = "Copied" OR status = "Already there" THEN LET(archived, archived + 1)
OP       RETURN(archived)
PERF     screen_off
ERR      standard
FIXTURE  grn_01.xlsx sheet=Grn cols=A:GRN No,B:Delivery Note File,C:Copy Status rows=10 data=A:values("GRN-1"x3,"GRN-2"x1,"GRN-3"x1,""x1,"..\x"x1,"GRN-4"x1,"GRN-5"x1,"GRN-6"x1)+boundary("GRN-1","GRN-2","GRN-3","","..\x","GRN-1","GRN-4","GRN-5","GRN-1","GRN-6");B:values("a.pdf"x2,"b.pdf"x1,"d.pdf"x1,"f.pdf"x1,"CON.pdf"x1,"c.pdf"x1,"e.pdf"x1,"g.pdf"x1,""x1)+boundary("a.pdf","b.pdf","d.pdf","a.pdf","f.pdf","CON.pdf","c.pdf","e.pdf","g.pdf","");C:values("old"x10)
FIXTURE  grn_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  grn_01.xlsx folder=Inbox files=a.pdf:100,b.pdf:80,c.pdf,e.pdf:60,g.pdf:40 locked=c.pdf
FIXTURE  grn_01.xlsx folder=Archive files=GRN-2\b.pdf:80,GRN-5\e.pdf:50
TEST     fixture=grn_01.xlsx call=ArchiveDeliveryNotes(Grn, Inbox, Archive) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Grn![Copy Status], "Already there") = 3 AND COUNTIF(Grn![Copy Status], "Missing") = 1 AND COUNTIF(Grn![Copy Status], "Bad name") = 3 AND COUNTIF(Grn![Copy Status], "Locked") = 1 AND COUNTIF(Grn![Copy Status], "Conflict") = 1 AND COUNTIF(Grn![Copy Status], "old") = 1 AND FILE_EXISTS("Archive\GRN-1\a.pdf") AND FILE_EXISTS("Archive\GRN-1\g.pdf") AND FILE_COUNT("Archive") = PRE_FILE_COUNT("Archive") + 2
TEST     fixture=grn_01.xlsx call=ArchiveDeliveryNotes(Grn, Inbox, Archive) assert=COUNTIF(Grn![Copy Status], "Copied") = 2 AND VALUE(Grn![Copy Status]2) = "Copied" AND VALUE(Grn![Copy Status]3) = "Already there" AND VALUE(Grn![Copy Status]4) = "Missing" AND VALUE(Grn![Copy Status]5) = "old" AND VALUE(Grn![Copy Status]6) = "Bad name" AND VALUE(Grn![Copy Status]8) = "Locked" AND VALUE(Grn![Copy Status]9) = "Conflict" AND VALUE(Grn![Copy Status]11) = "Bad name" AND FILE_COUNT("Inbox") = PRE_FILE_COUNT("Inbox") AND COUNTA(Grn![GRN No]) = PRE_COUNTA(Grn![GRN No])
TEST     fixture=grn_01.xlsx call=ArchiveDeliveryNotes(Notes, Inbox, Archive) assert=RESULT = -1 AND ERROR_COUNT = 1 AND FILE_COUNT("Archive") = PRE_FILE_COUNT("Archive")
