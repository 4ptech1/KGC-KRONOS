VCL 0.6
TASK
# ASSUME: Material Certs headers in row 1 include "Heat No", "Supplier" and "Grade"
# ASSUME: Supplier Grades already exists with its headers (Supplier, Grade, Certificates) and is rebuilt on every run
# ASSUME: supplier and grade are matched with spaces cleaned and case ignored; the first spelling found is shown
# ASSUME: rows with a blank supplier or a blank grade are ignored; the sheet is sorted by supplier A to Z
# ASSUME: every row with a supplier and a grade counts as one certificate; the macro returns the number of rows written
GOAL     Rebuild the Supplier Grades sheet with one row for each supplier and grade found together in Material Certs and the number of certificates, suppliers A to Z.
PROC     BuildSupplierGrades
KIND     Function
MODULE   modCerts
ARGS     certs As Worksheet, out As Worksheet
RET      Long
VAR      cSup As Long
VAR      cGrade As Long
VAR      key As String
VAR      counts As Dictionary
VAR      sups As Dictionary
VAR      grades As Dictionary
VAR      got As Long
VAR      s As String
VAR      g As String
VAR      outRow As Long
PRE      COL($certs, "Heat No") > 0; COL($certs, "Supplier") > 0; COL($certs, "Grade") > 0
MODIFIES $out
OP       LET(cSup, COL($certs, "Supplier"))
OP       LET(cGrade, COL($certs, "Grade"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Supplier", "Grade", "Certificates")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($certs, cSup) var=r
OP         SKIP_IF IS_EMPTY(CELL($certs, r, cSup)) OR IS_EMPTY(CELL($certs, r, cGrade))
OP         LET(s, CLEAN_TEXT(CELL($certs, r, cSup)))
OP         LET(g, CLEAN_TEXT(CELL($certs, r, cGrade)))
OP         LET(key, s & "|" & g)
OP         DICT_ADD(counts, key, 1, mode=sum)
OP         DICT_ADD(sups, key, s, mode=keep)
OP         DICT_ADD(grades, key, g, mode=keep)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=got)
OP         DICT_GET(sups, k, into=s)
OP         DICT_GET(grades, k, into=g)
OP         WRITE_ROW($out, outRow, s, g, got)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       RETURN(outRow - 2)
PERF     screen_off
ERR      standard
FIXTURE  certs_01.xlsx sheet=Certs cols=A:Heat No,B:Supplier,C:Grade rows=10 data=A:seq(1,1);B:values("Cora"x2,"Acme"x3,"Beta"x3,""x1,"BETA"x1)+boundary("Cora","Acme"," acme","Beta","","Acme","BETA","Acme","Beta","Cora");C:values("A516"x5,"A105"x3,""x1,"a516"x1)+boundary("A516","A516","a516","A105","A516","","  A105 ","A105","A105","A516")
FIXTURE  certs_01.xlsx sheet=SupplierGrades cols=A:Supplier,B:Grade,C:Certificates rows=0 stale=3
TEST     fixture=certs_01.xlsx call=BuildSupplierGrades(Certs, SupplierGrades) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND ROWS_USED(SupplierGrades) = 5 AND COUNTIF(SupplierGrades![Supplier], "stale") = 0 AND VALUE(SupplierGrades![Supplier]2) = "Acme" AND VALUE(SupplierGrades![Supplier]4) = "Beta" AND VALUE(SupplierGrades![Supplier]5) = "Cora" AND SUM(SupplierGrades![Certificates]) = 8
TEST     fixture=certs_01.xlsx call=BuildSupplierGrades(Certs, SupplierGrades) assert=SUMIFS(SupplierGrades![Certificates], SupplierGrades![Supplier], "Acme", SupplierGrades![Grade], "A516") = 2 AND SUMIFS(SupplierGrades![Certificates], SupplierGrades![Supplier], "Acme", SupplierGrades![Grade], "A105") = 1 AND SUMIFS(SupplierGrades![Certificates], SupplierGrades![Supplier], "Beta", SupplierGrades![Grade], "A105") = 3 AND SUMIFS(SupplierGrades![Certificates], SupplierGrades![Supplier], "Cora", SupplierGrades![Grade], "A516") = 2 AND COUNTA(Certs![Supplier]) = PRE_COUNTA(Certs![Supplier]) AND COUNTA(Certs![Grade]) = PRE_COUNTA(Certs![Grade])
TEST     fixture=certs_01.xlsx call=BuildSupplierGrades(SupplierGrades, Certs) assert=RESULT = -1 AND ERROR_COUNT = 1
