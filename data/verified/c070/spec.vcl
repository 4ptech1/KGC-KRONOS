VCL 0.6
TASK
# ASSUME: Material Certs headers in row 1 include "Supplier"; the Suppliers sheet already exists and is passed as out; it is cleared and rebuilt on every run
# ASSUME: suppliers are matched after CLEAN_TEXT (extra spaces removed), ignoring case, so 'Acme Steel' and 'ACME STEEL' are one supplier; the first spelling found is the one listed
# ASSUME: rows with no supplier are ignored; the list has the single heading Supplier in A1 and is sorted A to Z
# ASSUME: the out sheet never has a Heat No heading, so a certificates sheet passed as out by mistake is left alone instead of being cleared
GOAL     Rebuild the Suppliers sheet with each supplier from the Material Certs sheet listed once, in alphabetical order, and return how many suppliers there are.
PROC     ListSuppliers
KIND     Function
MODULE   modCerts
ARGS     certs As Worksheet, out As Worksheet
RET      Long
VAR      cSupplier As Long
VAR      supplier As String
VAR      seen As Dictionary
VAR      shown As String
VAR      outRow As Long
PRE      COL($certs, "Supplier") > 0; COL($out, "Heat No") = 0
MODIFIES $out
OP       LET(cSupplier, COL($certs, "Supplier"))
OP       LOOP_ROWS dir=down from=2 to=LASTROW($certs, cSupplier) var=r
OP         SKIP_IF IS_EMPTY(CELL($certs, r, cSupplier))
OP         LET(supplier, CLEAN_TEXT(CELL($certs, r, cSupplier)))
OP         DICT_ADD(seen, supplier, supplier, mode=keep)
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Supplier")
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(seen)
OP         DICT_GET(seen, k, into=shown)
OP         WRITE_ROW($out, outRow, shown)
OP         LET(outRow, outRow + 1)
OP       SORT(USED_RANGE($out), key=1, order=asc, header=yes)
OP       RETURN(COUNT(seen))
PERF     screen_off calc_manual
ERR      standard
FIXTURE  certs_01.xlsx sheet=Certs cols=A:Heat No,B:MTC No,C:Supplier rows=15 data=A:seq(1,1);B:text(8);C:values("Zenith Alloys"x1,"Acme  Steel"x1,"ACME STEEL"x1,""x2,"bohler plate"x1," Acme Steel "x1,"Nippon Plate"x1,"acme steel"x1,"zenith alloys"x1,"Baosteel"x2,"NIPPON PLATE"x1,"Delta Forge"x1,"  "x1)+boundary("Zenith Alloys","Acme  Steel","ACME STEEL","","bohler plate"," Acme Steel ","Nippon Plate","acme steel","","zenith alloys","Baosteel","NIPPON PLATE","Delta Forge","Baosteel","  ")
FIXTURE  certs_01.xlsx sheet=Suppliers cols=A:Supplier rows=0 stale=3
FIXTURE  certs_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  certs_01.xlsx sheet=OldCerts cols=A:Heat No,B:MTC No,C:Supplier rows=3 stale=1 data=A:seq(1,1);B:text(8);C:text(8)
FIXTURE  certs_02.xlsx sheet=Certs cols=A:Heat No,B:MTC No,C:Supplier rows=3 data=A:seq(1,1);B:text(8);C:values(""x3)
FIXTURE  certs_02.xlsx sheet=Suppliers cols=A:Supplier rows=0 stale=2
TEST     fixture=certs_01.xlsx call=ListSuppliers(Certs, Suppliers) runs=2 assert=RESULT = 6 AND ERROR_COUNT = 0 AND COUNTA(Suppliers![Supplier]) = 7 AND ROWS_USED(Suppliers) = 7 AND VALUE(Suppliers![Supplier]1) = "Supplier" AND VALUE(Suppliers![Supplier]2) = "Acme Steel" AND VALUE(Suppliers![Supplier]3) = "Baosteel" AND VALUE(Suppliers![Supplier]4) = "bohler plate" AND VALUE(Suppliers![Supplier]5) = "Delta Forge" AND VALUE(Suppliers![Supplier]6) = "Nippon Plate" AND VALUE(Suppliers![Supplier]7) = "Zenith Alloys" AND COUNTIF(Suppliers![Supplier], "stale") = 0 AND FILL_COUNT(Suppliers![Supplier], #FF0000) = 0
TEST     fixture=certs_01.xlsx call=ListSuppliers(Certs, Suppliers) assert=RESULT = 6 AND COUNTA(Certs![Supplier]) = PRE_COUNTA(Certs![Supplier]) AND COUNTA(Certs![Heat No]) = PRE_COUNTA(Certs![Heat No]) AND COUNTA(Certs![MTC No]) = PRE_COUNTA(Certs![MTC No])
TEST     fixture=certs_02.xlsx call=ListSuppliers(Certs, Suppliers) assert=RESULT = 0 AND ERROR_COUNT = 0 AND ROWS_USED(Suppliers) = 1 AND VALUE(Suppliers![Supplier]1) = "Supplier" AND COUNTIF(Suppliers![Supplier], "stale") = 0
TEST     fixture=certs_01.xlsx call=ListSuppliers(Notes, Suppliers) assert=RESULT = -1 AND ERROR_COUNT = 1
TEST     fixture=certs_01.xlsx call=ListSuppliers(Certs, OldCerts) assert=RESULT = -1 AND ERROR_COUNT = 1 AND COUNTA(OldCerts![Supplier]) = PRE_COUNTA(OldCerts![Supplier])
