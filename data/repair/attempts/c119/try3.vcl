VCL 0.6
TASK
# ASSUME: Material Certs headers in row 1 include "Supplier" and "Grade"; the real sheets are "Material Certs" and "Supplier Grades"
# ASSUME: the Supplier Grades sheet exists and is rebuilt on every run; each row of Material Certs is one certificate
# ASSUME: supplier and grade are matched with CLEAN_TEXT, ignoring case; the text shown is the first one met, cleaned
# ASSUME: rows are sorted by supplier A to Z, then by grade A to Z
GOAL     Rebuild the Supplier Grades sheet with one row per supplier and grade pair and its number of certificates, suppliers A to Z.
PROC     SummariseSupplierGrades
KIND     Function
MODULE   modCerts
ARGS     certs As Worksheet, out As Worksheet
RET      Long
VAR      cSupplier As Long
VAR      cGrade As Long
VAR      rowOf As Dictionary
VAR      tally As Dictionary
VAR      pairKey As String
VAR      at As Long
VAR      n As Long
VAR      outRow As Long
PRE      COL($certs, "Supplier") > 0; COL($certs, "Grade") > 0
MODIFIES $out
OP       LET(cSupplier, COL($certs, "Supplier"))
OP       LET(cGrade, COL($certs, "Grade"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Supplier", "Grade", "Certificates")
OP       LET(outRow, 2)
OP       LOOP_ROWS dir=down from=2 to=LASTROW($certs, cSupplier) var=r
OP         SKIP_IF IS_EMPTY(CELL($certs, r, cSupplier))
OP         SKIP_IF IS_EMPTY(CELL($certs, r, cGrade))
OP         LET(pairKey, CLEAN_TEXT(CELL($certs, r, cSupplier)) & "|" & CLEAN_TEXT(CELL($certs, r, cGrade)))
OP         IF NOT EXISTS(rowOf, pairKey)
OP           WRITE_ROW($out, outRow, CLEAN_TEXT(CELL($certs, r, cSupplier)), CLEAN_TEXT(CELL($certs, r, cGrade)), 0)
OP           DICT_ADD(rowOf, pairKey, outRow, mode=keep)
OP           LET(outRow, outRow + 1)
OP         DICT_ADD(tally, pairKey, 1, mode=sum)
OP       FOR_EACH var=k in=KEYS(rowOf)
OP         DICT_GET(rowOf, k, into=at)
OP         DICT_GET(tally, k, into=n)
OP         SET_VALUE(CELL($out, at, 3), n)
OP       SORT(USED_RANGE($out), key=2, order=asc, header=yes)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       RETURN(COUNT(rowOf))
PERF     screen_off calc_manual
ERR      standard
FIXTURE  certs_01.xlsx sheet=Certs cols=A:Heat No,B:Supplier,C:Grade rows=12 stale=2 data=A:seq(1,1);B:values("Zeta Alloys"x2,"Acme Steel"x3," ACME Steel "x1,"acme steel"x1,"Beta Forge"x3,""x1,"  "x1)+boundary("Zeta Alloys","Acme Steel"," ACME Steel ","Beta Forge","Acme Steel","","Beta Forge","  ","acme steel","Beta Forge","Zeta Alloys","Acme Steel");C:values("SA-516"x5,"sa-516 "x1,"SA-105"x5,""x1)+boundary("SA-516","SA-516","sa-516 ","SA-105","SA-105","SA-105","","SA-516","SA-516","SA-105","SA-516","SA-105")
FIXTURE  certs_01.xlsx sheet=Grades cols=A:Supplier,B:Grade,C:Certificates rows=0 stale=3
TEST     fixture=certs_01.xlsx call=SummariseSupplierGrades(Certs, Grades) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND ROWS_USED(Grades) = 5 AND VALUE(Grades![Supplier]2) = "Acme Steel" AND VALUE(Grades![Grade]2) = "SA-105" AND VALUE(Grades![Certificates]2) = 2 AND VALUE(Grades![Supplier]3) = "Acme Steel" AND VALUE(Grades![Grade]3) = "SA-516" AND VALUE(Grades![Certificates]3) = 3 AND VALUE(Grades![Supplier]4) = "Beta Forge" AND VALUE(Grades![Certificates]4) = 2 AND VALUE(Grades![Supplier]5) = "Zeta Alloys" AND VALUE(Grades![Certificates]5) = 2 AND COUNTIF(Grades![Supplier], "stale") = 0
TEST     fixture=certs_01.xlsx call=SummariseSupplierGrades(Certs, Grades) assert=SUM(Grades![Certificates]) = 9 AND COUNTA(Certs![Heat No]) = PRE_COUNTA(Certs![Heat No]) AND COUNTA(Certs![Supplier]) = PRE_COUNTA(Certs![Supplier]) AND COUNTIF(Certs![Grade], "sa-516") = PRE_COUNTIF(Certs![Grade], "sa-516")
TEST     fixture=certs_01.xlsx call=SummariseSupplierGrades(Grades, Certs) assert=RESULT = -1 AND ERROR_COUNT = 1
