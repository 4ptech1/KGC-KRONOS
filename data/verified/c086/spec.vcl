VCL 0.6
TASK
# ASSUME: Consumable Stock headers in row 1 include "Batch No", "Cert File" and "Cert Status"
# ASSUME: every row has a Batch No; rows are checked from row 2 down to the last Batch No
# ASSUME: Cert File holds the file's name with its extension, not a path; it matches ignoring case and spaces at the ends
# ASSUME: only files directly in the Batch Certs folder count; a file in a folder below it, or a hidden file, is Missing
# ASSUME: a Cert File that holds only spaces counts as no file name; Cert Status is rewritten on every row on every run
# ASSUME: the macro only lists the folder's file names: no file is opened, changed, copied or deleted
GOAL     Mark each Consumable Stock row Found or Missing by whether its Cert File is in the Batch Certs folder, or No file name when it has none, and return how many are Missing.
PROC     CheckBatchCertFiles
KIND     Function
MODULE   modCerts
ARGS     ws As Worksheet, certs As String
RET      Long
VAR      cBatch As Long
VAR      cFile As Long
VAR      cStatus As Long
VAR      files As Collection
VAR      status As String
VAR      n As Long
PRE      COL($ws, "Batch No") > 0; COL($ws, "Cert File") > 0; COL($ws, "Cert Status") > 0; FOLDER_EXISTS($certs)
MODIFIES $ws["Cert Status"]
OP       LET(cBatch, COL($ws, "Batch No"))
OP       LET(cFile, COL($ws, "Cert File"))
OP       LET(cStatus, COL($ws, "Cert Status"))
OP       LIST_FILES($certs, "*", into=files)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cBatch) var=r
OP         IF IS_EMPTY(CELL($ws, r, cFile))
OP           SET_VALUE(CELL($ws, r, cStatus), "No file name")
OP         ELSE
OP           LET(status, "Missing")
OP           FOR_EACH var=f in=ITEMS(files)
OP             IF SAME(f, CELL($ws, r, cFile)) THEN LET(status, "Found")
OP           SET_VALUE(CELL($ws, r, cStatus), status)
OP           IF status = "Missing" THEN LET(n, n + 1)
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  certs_01.xlsx sheet=Stock cols=A:Batch No,B:Description,C:Cert File,D:Cert Status rows=9 data=A:seq(101,1);B:text(10);C:values("b101.pdf"x1," B102.PDF "x1,"b103.pdf"x1,""x1,"b105.pdf"x1,"b106.pdf"x1,"   "x1,"nofile.pdf"x1,"z-last.pdf"x1)+boundary("b101.pdf"," B102.PDF ","b103.pdf","","b105.pdf","b106.pdf","   ","nofile.pdf","z-last.pdf");D:values("Missing"x2,""x2,"Found"x3,"old"x2)+boundary("Missing","","Found","Found","old","","Missing","Found","old")
FIXTURE  certs_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  certs_01.xlsx folder=Certs files=b101.pdf:100,b102.pdf:80,b105.pdf:60,old\b106.pdf:50,notes.txt:30,z-last.pdf:40 locked=b105.pdf
TEST     fixture=certs_01.xlsx call=CheckBatchCertFiles(Stock, Certs) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(Stock![Cert Status], "Found") = 4 AND COUNTIF(Stock![Cert Status], "Missing") = 3 AND COUNTIF(Stock![Cert Status], "No file name") = 2 AND VALUE(Stock![Cert Status]2) = "Found" AND VALUE(Stock![Cert Status]3) = "Found" AND VALUE(Stock![Cert Status]4) = "Missing" AND VALUE(Stock![Cert Status]5) = "No file name" AND VALUE(Stock![Cert Status]6) = "Found" AND VALUE(Stock![Cert Status]7) = "Missing" AND VALUE(Stock![Cert Status]8) = "No file name" AND VALUE(Stock![Cert Status]9) = "Missing" AND VALUE(Stock![Cert Status]10) = "Found" AND FILE_COUNT("Certs") = PRE_FILE_COUNT("Certs")
TEST     fixture=certs_01.xlsx call=CheckBatchCertFiles(Stock, Certs) assert=RESULT = 3 AND COUNTA(Stock![Cert File]) = PRE_COUNTA(Stock![Cert File]) AND COUNTA(Stock![Batch No]) = PRE_COUNTA(Stock![Batch No]) AND COUNTA(Stock![Cert Status]) = 10 AND FILE_COUNT("Certs") = PRE_FILE_COUNT("Certs") AND FILE_EXISTS("Certs\b105.pdf") AND FILE_EXISTS("Certs\old\b106.pdf")
TEST     fixture=certs_01.xlsx call=CheckBatchCertFiles(Notes, Certs) assert=RESULT = -1 AND ERROR_COUNT = 1 AND FILE_COUNT("Certs") = PRE_FILE_COUNT("Certs")
