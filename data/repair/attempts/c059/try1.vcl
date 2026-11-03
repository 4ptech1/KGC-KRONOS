VCL 0.6
TASK
# ASSUME: Material Certs headers in row 1 include "Heat No", "MTC No" and "Supplier"; the real sheets are "Material Certs" and "Supplier Summary", the TESTs use MaterialCerts and SupplierSummary
# ASSUME: the Supplier Summary sheet already exists and is passed in as out; it is rebuilt from scratch on every run
# ASSUME: every row with a supplier is one certificate, and it has no MTC No when that cell is blank or only spaces
# ASSUME: suppliers are matched ignoring case and spaces at the ends and are listed as first typed; rows with no supplier are ignored
# ASSUME: the summary has the headings Supplier, Certificates and No MTC No; suppliers with the same count may come in any order
GOAL     Rebuild the Supplier Summary sheet with the number of certificates per supplier and how many of them have no MTC No, the supplier with the most certificates first.
PROC     SummariseCertsBySupplier
KIND     Function
MODULE   modCerts
ARGS     certs As Worksheet, out As Worksheet
RET      Long
VAR      cMtc As Long
VAR      cSupplier As Long
VAR      sup As String
VAR      counts As Dictionary
VAR      noMtc As Dictionary
VAR      total As Long
VAR      noMtcCount As Long
VAR      outRow As Long
PRE      COL($certs, "MTC No") > 0; COL($certs, "Supplier") > 0
MODIFIES $out
OP       LET(cMtc, COL($certs, "MTC No"))
OP       LET(cSupplier, COL($certs, "Supplier"))
OP       CLEAR(USED_RANGE($out), what=all)
OP       WRITE_ROW($out, 1, "Supplier", "Certificates", "No MTC No")
OP       LOOP_ROWS dir=down from=2 to=LASTROW($certs, cSupplier) var=r
OP         SKIP_IF IS_EMPTY(CELL($certs, r, cSupplier))
OP         LET(sup, TRIM(CELL($certs, r, cSupplier)))
OP         DICT_ADD(counts, sup, 1, mode=sum)
OP         DICT_ADD(noMtc, sup, IIF(IS_EMPTY(CELL($certs, r, cMtc)), 1, 0), mode=sum)
OP       LET(outRow, 2)
OP       FOR_EACH var=k in=KEYS(counts)
OP         DICT_GET(counts, k, into=total)
OP         DICT_GET(noMtc, k, into=noMtcCount)
OP         WRITE_ROW($out, outRow, k, total, noMtcCount)
OP         LET(outRow, outRow + 1)
OP       IF COUNT(counts) > 1
OP         SORT(USED_RANGE($out), key=2, order=desc, header=yes)
OP       RETURN(COUNT(counts))
PERF     screen_off calc_manual
ERR      standard
FIXTURE  material_certs_01.xlsx sheet=MaterialCerts cols=A:Heat No,B:MTC No,C:Supplier rows=14 data=A:seq(501,1);B:text(6)+boundary("","M-102","","M-104","M-105","","M-107","M-108","","M-110","M-111","","M-113"," ");C:values("Apex Steel"x3," Apex Steel "x1,"apex steel"x1,"Borealis Plate"x4,"Cobalt Forge"x2,"Delta Alloys"x1,""x2)+boundary("Delta Alloys","Cobalt Forge","Borealis Plate","Apex Steel","","Apex Steel","Borealis Plate"," Apex Steel ","Cobalt Forge","Borealis Plate","Apex Steel","","Borealis Plate","apex steel")
FIXTURE  material_certs_01.xlsx sheet=SupplierSummary cols=A:Supplier,B:Certificates,C:No MTC No rows=0 stale=6
TEST     fixture=material_certs_01.xlsx call=SummariseCertsBySupplier(MaterialCerts, SupplierSummary) runs=2 assert=RESULT = 4 AND ERROR_COUNT = 0 AND ROWS_USED(SupplierSummary) = 5 AND COUNTIF(SupplierSummary![Supplier], "stale") = 0 AND FILL_COUNT(SupplierSummary![Supplier], #FF0000) = 0 AND VALUE(SupplierSummary![Supplier]2) = "Apex Steel" AND VALUE(SupplierSummary![Certificates]2) = 5 AND VALUE(SupplierSummary![No MTC No]2) = 2
TEST     fixture=material_certs_01.xlsx call=SummariseCertsBySupplier(MaterialCerts, SupplierSummary) assert=VALUE(SupplierSummary![Supplier]3) = "Borealis Plate" AND VALUE(SupplierSummary![Certificates]3) = 4 AND VALUE(SupplierSummary![No MTC No]3) = 1 AND VALUE(SupplierSummary![Supplier]4) = "Cobalt Forge" AND VALUE(SupplierSummary![Certificates]4) = 2 AND VALUE(SupplierSummary![No MTC No]4) = 1 AND VALUE(SupplierSummary![Supplier]5) = "Delta Alloys" AND VALUE(SupplierSummary![Certificates]5) = 1 AND VALUE(SupplierSummary![No MTC No]5) = 1 AND SUM(SupplierSummary![Certificates]) = 12 AND COUNTA(MaterialCerts![Heat No]) = PRE_COUNTA(MaterialCerts![Heat No]) AND COUNTA(MaterialCerts![Supplier]) = PRE_COUNTA(MaterialCerts![Supplier])
TEST     fixture=material_certs_01.xlsx call=SummariseCertsBySupplier(SupplierSummary, MaterialCerts) assert=RESULT = -1 AND ERROR_COUNT = 1
