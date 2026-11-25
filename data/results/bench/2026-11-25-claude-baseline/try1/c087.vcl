VCL 0.7
TASK
# ASSUME: the Certs folder holds heat.txt with one heat number; leading and trailing spaces and a final line break are ignored
# ASSUME: Material Certs headers in row 1 include "Heat No" and "Remark"; Heat No is matched with SAME and only the first match is marked
# ASSUME: when heat.txt cannot be read (not there, locked, not text) a message says so, nothing changes and 0 is returned; 0 also means the heat is not on the sheet
GOAL     Read the heat number from heat.txt in the Certs folder, write Verified in the Remark of that heat on the Material Certs sheet and return its row.
PROC     VerifyHeatFromFile
KIND     Function
MODULE   modCerts
ARGS     ws As Worksheet, certs As String
RET      Long
VAR      cHeat As Long
VAR      cRemark As Long
VAR      text As String
VAR      status As String
VAR      want As String
VAR      foundRow As Long
PRE      COL($ws, "Heat No") > 0; COL($ws, "Remark") > 0; FOLDER_EXISTS($certs)
MODIFIES $ws["Remark"]
OP       LET(cHeat, COL($ws, "Heat No"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       READ_TEXT(PATH_JOIN($certs, "heat.txt"), text, into=status)
OP       IF status <> "Read" THEN MSG("heat.txt could not be read: " & status)
OP       EXIT_IF status <> "Read"
OP       LET(want, TRIM(text))
OP       EXIT_IF want = ""
OP       LOOP_ROWS dir=down from=2 to=LASTROW($ws, cHeat) var=r
OP         IF SAME(CELL($ws, r, cHeat), want)
OP           SET_VALUE(CELL($ws, r, cRemark), "Verified")
OP           LET(foundRow, r)
OP         EXIT_LOOP_IF foundRow > 0
OP       RETURN(foundRow)
PERF     screen_off
ERR      standard
FIXTURE  certs_01.xlsx sheet=MaterialCerts cols=A:Heat No,B:MTC No,C:Supplier,D:Remark rows=5 data=A:values("H-1001"x1,"H-1002"x1,"H-1003"x1,"H-1004"x1,"H-1005"x1)+boundary("H-1001","H-1002","H-1003","H-1004","H-1005");B:text(6);C:text(8);D:values("old"x5)
FIXTURE  certs_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  certs_01.xlsx folder=Certs files=heat.txt:"H-1002",extra.txt:20 locked=extra.txt
FIXTURE  certs_02.xlsx sheet=MaterialCerts cols=A:Heat No,B:MTC No,C:Supplier,D:Remark rows=5 data=A:values("H-1001"x1,"H-1002"x1,"H-1003"x1,"H-1004"x1,"H-1005"x1)+boundary("H-1001","H-1002","H-1003","H-1004","H-1005");B:text(6);C:text(8);D:values("old"x5)
FIXTURE  certs_02.xlsx folder=Certs files=other.txt:30
FIXTURE  certs_03.xlsx sheet=MaterialCerts cols=A:Heat No,B:MTC No,C:Supplier,D:Remark rows=5 data=A:values("H-1001"x1,"H-1002"x1,"H-1003"x1,"H-1004"x1,"H-1005"x1)+boundary("H-1001","H-1002","H-1003","H-1004","H-1005");B:text(6);C:text(8);D:values("old"x5)
FIXTURE  certs_03.xlsx folder=Certs files=heat.txt:"H-9999"
FIXTURE  certs_04.xlsx sheet=MaterialCerts cols=A:Heat No,B:MTC No,C:Supplier,D:Remark rows=5 data=A:values("H-1001"x1,"H-1002"x1,"H-1003"x1,"H-1004"x1,"H-1005"x1)+boundary("H-1001","H-1002","H-1003","H-1004","H-1005");B:text(6);C:text(8);D:values("old"x5)
FIXTURE  certs_04.xlsx folder=Certs files=heat.txt:"H-1002" locked=heat.txt
TEST     fixture=certs_01.xlsx call=VerifyHeatFromFile(MaterialCerts, Certs) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND COUNTIF(MaterialCerts![Remark], "Verified") = 1 AND COUNTIF(MaterialCerts![Remark], "old") = 4 AND VALUE(MaterialCerts![Remark]3) = "Verified" AND FILE_TEXT("Certs\heat.txt") = "H-1002"
TEST     fixture=certs_01.xlsx call=VerifyHeatFromFile(MaterialCerts, Certs) assert=COUNTA(MaterialCerts![Heat No]) = PRE_COUNTA(MaterialCerts![Heat No]) AND FILE_COUNT("Certs") = PRE_FILE_COUNT("Certs")
TEST     fixture=certs_02.xlsx call=VerifyHeatFromFile(MaterialCerts, Certs) assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(MaterialCerts![Remark], "Verified") = 0 AND COUNTIF(MaterialCerts![Remark], "old") = 5
TEST     fixture=certs_03.xlsx call=VerifyHeatFromFile(MaterialCerts, Certs) assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(MaterialCerts![Remark], "Verified") = 0 AND COUNTIF(MaterialCerts![Remark], "old") = 5
TEST     fixture=certs_04.xlsx call=VerifyHeatFromFile(MaterialCerts, Certs) assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(MaterialCerts![Remark], "old") = 5
TEST     fixture=certs_01.xlsx call=VerifyHeatFromFile(Notes, Certs) assert=RESULT = -1 AND ERROR_COUNT = 1
