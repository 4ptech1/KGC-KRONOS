VCL 0.6
TASK
# ASSUME: File List is an existing sheet; it is rebuilt from scratch on every run
# ASSUME: the Dispatch Docs folder is given as a full path and only the files directly in it are listed, not those in sub-folders
# ASSUME: row 1 holds the heading "File Name" and the PDF names, in name order ignoring case, start in row 2 of column A
# ASSUME: a file counts as a PDF when its whole name ends in .pdf, in any case
# ASSUME: the count is shown in a message and also returned
GOAL     List the PDF files in the Dispatch Docs folder on the File List sheet, one per row in name order, and tell the user how many there are.
PROC     ListDispatchPdfs
KIND     Function
MODULE   modDispatch
ARGS     out As Worksheet, docs As String
RET      Long
VAR      pdfs As Collection
VAR      n As Long
PRE      FOLDER_EXISTS($docs)
MODIFIES $out
OP       LIST_FILES($docs, "*.pdf", into=pdfs)
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "File Name")
OP       FOR_EACH var=f in=ITEMS(pdfs)
OP         LET(n, n + 1)
OP         WRITE_ROW($out, n + 1, f)
OP       MSG("PDF files listed: " & n)
OP       RETURN(n)
PERF     screen_off
ERR      standard
FIXTURE  dispatch_01.xlsx sheet=FileList cols=A:File Name rows=0 stale=3
FIXTURE  dispatch_01.xlsx folder=Docs files=Zed.pdf:30,alpha.pdf:10,Beta.PDF:20,gamma.pdf:40,notes.txt:5,scan.pdf.bak:5,old\ancient.pdf:50 locked=gamma.pdf
FIXTURE  dispatch_01.xlsx folder=Texts files=readme.txt:5
TEST     fixture=dispatch_01.xlsx call=ListDispatchPdfs(FileList, Docs) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND ROWS_USED(FileList) = 5 AND COUNTIF(FileList![File Name], "stale") = 0 AND VALUE(FileList![File Name]1) = "File Name" AND VALUE(FileList![File Name]2) = "alpha.pdf" AND VALUE(FileList![File Name]3) = "Beta.PDF" AND VALUE(FileList![File Name]4) = "gamma.pdf" AND VALUE(FileList![File Name]5) = "Zed.pdf"
TEST     fixture=dispatch_01.xlsx call=ListDispatchPdfs(FileList, Docs) assert=ERROR_COUNT = 0 AND FILE_COUNT("Docs") = PRE_FILE_COUNT("Docs") AND COUNTIF(FileList![File Name], "*.txt") = 0 AND COUNTIF(FileList![File Name], "*.bak") = 0 AND COUNTIF(FileList![File Name], "ancient.pdf") = 0
TEST     fixture=dispatch_01.xlsx call=ListDispatchPdfs(FileList, Texts) assert=RESULT = 0 AND ERROR_COUNT = 0 AND ROWS_USED(FileList) = 1 AND COUNTIF(FileList![File Name], "stale") = 0 AND FILE_COUNT("Texts") = PRE_FILE_COUNT("Texts")
TEST     fixture=dispatch_01.xlsx call=ListDispatchPdfs(FileList, "C:\Nowhere") assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTIF(FileList![File Name], "stale") = 3
