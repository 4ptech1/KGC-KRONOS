VCL 0.6
TASK
# ASSUME: Calibration Tracker headers in row 1 include "Instrument ID" and "Cert Status"
# ASSUME: a certificate is a PDF directly in the Cal Certs folder named after the trimmed instrument ID, such as PG-104.pdf; case is ignored and subfolders are not searched
# ASSUME: a row with a blank Instrument ID is left as it is; an ID that cannot be a file name simply has no PDF, so it is Missing
# ASSUME: no file is created, changed or deleted; the macro only reads the folder's list of names
GOAL     Mark each Calibration Tracker row Found or Missing in Cert Status by whether its instrument has a PDF certificate in the Cal Certs folder, and return how many are Missing.
PROC     CheckCertFiles
KIND     Function
MODULE   modCerts
ARGS     tracker As Worksheet, certs As String
RET      Long
VAR      cInst As Long
VAR      cStatus As Long
VAR      pdfs As Collection
VAR      have As Dictionary
VAR      missing As Long
PRE      COL($tracker, "Instrument ID") > 0; COL($tracker, "Cert Status") > 0; FOLDER_EXISTS($certs)
MODIFIES $tracker["Cert Status"]
OP       LET(cInst, COL($tracker, "Instrument ID"))
OP       LET(cStatus, COL($tracker, "Cert Status"))
OP       LIST_FILES($certs, "*.pdf", into=pdfs)
OP       FOR_EACH var=f in=ITEMS(pdfs)
OP         DICT_ADD(have, f, 1, mode=keep)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($tracker, cInst) var=r
OP         SKIP_IF IS_EMPTY(CELL($tracker, r, cInst))
OP         IF EXISTS(have, TRIM(CELL($tracker, r, cInst)) & ".pdf")
OP           SET_VALUE(CELL($tracker, r, cStatus), "Found")
OP         ELSE
OP           SET_VALUE(CELL($tracker, r, cStatus), "Missing")
OP           LET(missing, missing + 1)
OP       RETURN(missing)
PERF     screen_off
ERR      standard
FIXTURE  certs_01.xlsx sheet=Tracker cols=A:Instrument ID,B:Cert Status rows=7 data=A:values("PG-104"x2," pg-106 "x1,"PG-105"x1,"PG-107"x1,""x1,"PG-108"x1)+boundary("PG-104"," pg-106 ","PG-105","PG-107","","PG-108","PG-104");B:values("old"x7)
FIXTURE  certs_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  certs_01.xlsx folder=Certs files=PG-104.pdf:100,PG-106.pdf:80,pg-108.PDF:50,PG-105.txt:30,PG-109.pdf:10,Old\PG-107.pdf:20 locked=PG-106.pdf
TEST     fixture=certs_01.xlsx call=CheckCertFiles(Tracker, Certs) runs=2 assert=RESULT = 2 AND ERROR_COUNT = 0 AND COUNTIF(Tracker![Cert Status], "Found") = 4 AND COUNTIF(Tracker![Cert Status], "Missing") = 2 AND COUNTIF(Tracker![Cert Status], "old") = 1
TEST     fixture=certs_01.xlsx call=CheckCertFiles(Tracker, Certs) assert=VALUE(Tracker![Cert Status]2) = "Found" AND VALUE(Tracker![Cert Status]3) = "Found" AND VALUE(Tracker![Cert Status]4) = "Missing" AND VALUE(Tracker![Cert Status]5) = "Missing" AND VALUE(Tracker![Cert Status]6) = "old" AND VALUE(Tracker![Cert Status]7) = "Found" AND VALUE(Tracker![Cert Status]8) = "Found" AND COUNTA(Tracker![Instrument ID]) = PRE_COUNTA(Tracker![Instrument ID]) AND FILE_COUNT("Certs") = PRE_FILE_COUNT("Certs") AND FILE_EXISTS("Certs\PG-104.pdf")
TEST     fixture=certs_01.xlsx call=CheckCertFiles(Notes, Certs) assert=RESULT = -1 AND ERROR_COUNT = 1 AND FILE_COUNT("Certs") = PRE_FILE_COUNT("Certs")
