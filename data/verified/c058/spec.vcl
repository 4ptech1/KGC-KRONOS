VCL 0.6
TASK
# ASSUME: Material Certs headers (row 1) include "Material Grade"
# ASSUME: each row is one certificate; a row with a blank Material Grade is skipped
# ASSUME: grades are compared ignoring case, with extra spaces cleaned; the first spelling seen is shown
# ASSUME: grades with the same number of certificates may come in any order
# ASSUME: Grade Summary is a new sheet after the last one; if it is already there the run changes nothing and returns 0
# ASSUME: the real sheet is "Material Certs"; the TESTs use Certs
GOAL     Make a Grade Summary sheet that lists each material grade once with its number of certificates, the biggest count first.
PROC     MakeGradeSummary
KIND     Function
MODULE   modCerts
ARGS     certs As Worksheet, wb As Workbook
RET      Long
VAR      cGrade As Long
VAR      grades As Dictionary
VAR      out As Worksheet
VAR      outRow As Long
VAR      got As Long
PRE      COL($certs, "Material Grade") > 0
MODIFIES $wb sheets
OP       IF SHEET_EXISTS($wb, "Grade Summary")
OP         LOG("Grade Summary is already there; left as it is")
OP         RETURN(0)
OP       LET(cGrade, COL($certs, "Material Grade"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($certs, cGrade) var=r
OP         SKIP_IF IS_EMPTY(CELL($certs, r, cGrade))
OP         DICT_ADD(grades, CLEAN_TEXT(CELL($certs, r, cGrade)), 1, mode=sum)
OP       ADD_SHEET($wb, name="Grade Summary", into=out)
OP       WRITE_ROW(out, 1, "Material Grade", "Certificates")
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(grades)
OP         DICT_GET(grades, k, into=got)
OP         WRITE_ROW(out, outRow, k, got)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE(out), key=2, order=desc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off
ERR      standard
FIXTURE  certs_01.xlsx sheet=Certs cols=A:Heat No,B:MTC No,C:Material Grade,D:Cert Date rows=17 data=A:seq(100,1);B:seq(5001,1);C:values("A105"x1,"304L"x2,"A516 60"x3,"A106 B"x3,"A106 B "x1,"A516 70"x3,"a516 70"x1,"A516  70"x1,""x2)+boundary("A105","304L","A516 60","A106 B","A516 70","","A516 70","a516 70","A106 B","A516 60","304L","A516  70","A106 B","A106 B ","A516 60","","A516 70");D:dates(2025-01-01,2026-09-30)
FIXTURE  certs_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
TEST     fixture=certs_01.xlsx call=MakeGradeSummary(Certs, certs_01.xlsx) assert=RESULT = 5 AND ERROR_COUNT = 0 AND ROWS_USED('Grade Summary') = 6 AND VALUE('Grade Summary'![Material Grade]2) = "A516 70" AND VALUE('Grade Summary'![Certificates]2) = 5 AND VALUE('Grade Summary'![Material Grade]3) = "A106 B" AND VALUE('Grade Summary'![Certificates]3) = 4 AND VALUE('Grade Summary'![Material Grade]4) = "A516 60" AND VALUE('Grade Summary'![Certificates]4) = 3 AND VALUE('Grade Summary'![Material Grade]5) = "304L" AND VALUE('Grade Summary'![Certificates]5) = 2 AND VALUE('Grade Summary'![Material Grade]6) = "A105" AND VALUE('Grade Summary'![Certificates]6) = 1
TEST     fixture=certs_01.xlsx call=MakeGradeSummary(Certs, certs_01.xlsx) runs=2 assert=ERROR_COUNT = 0 AND SHEET_EXISTS("Grade Summary") AND COUNTA('Grade Summary'![Material Grade]) = 6 AND SUM('Grade Summary'![Certificates]) = 15
TEST     fixture=certs_01.xlsx call=MakeGradeSummary(Certs, certs_01.xlsx) assert=COUNTA(Certs![Material Grade]) = PRE_COUNTA(Certs![Material Grade]) AND COUNTA(Certs![MTC No]) = PRE_COUNTA(Certs![MTC No]) AND COUNTA(Certs![Cert Date]) = PRE_COUNTA(Certs![Cert Date])
TEST     fixture=certs_01.xlsx call=MakeGradeSummary(Notes, certs_01.xlsx) assert=RESULT = -1 AND ERROR_COUNT = 1 AND SHEET_EXISTS("Grade Summary") = 0
