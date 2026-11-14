VCL 0.6
TASK
# ASSUME: Material Certs headers in row 1 include "Heat No" and "Remark"; the Certs folder exists
# ASSUME: heat.txt holds just the heat number on one line; it matches a Heat No ignoring case and spaces at the ends
# ASSUME: when the heat is on several rows only the first one gets Verified and its row is returned; every other Remark is left as it is
# ASSUME: when heat.txt is missing, locked, not plain text or holds no heat number, a message says so, nothing changes and the result is -1
# ASSUME: heat.txt is only read: no file is written, changed, copied or deleted
GOAL     Read the heat number in heat.txt in the Certs folder, write Verified in the Remark of the first Material Certs row with that Heat No and return its row, or 0 when it is not there.
PROC     VerifyHeatFromFile
KIND     Function
MODULE   modCerts
ARGS     ws As Worksheet, certs As String
RET      Long
VAR      cHeat As Long
VAR      cRemark As Long
VAR      heat As String
VAR      status As String
VAR      found As Long
PRE      COL($ws, "Heat No") > 0; COL($ws, "Remark") > 0; FOLDER_EXISTS($certs)
MODIFIES $ws["Remark"]
OP       LET(cHeat, COL($ws, "Heat No"))
OP       LET(cRemark, COL($ws, "Remark"))
OP       READ_TEXT(PATH_JOIN($certs, "heat.txt"), heat, into=status)
OP       IF status <> "Read"
OP         MSG("heat.txt was not read (" & status & "), so nothing was changed.")
OP         LET(found, -1)
OP       ELSE
OP         IF IS_EMPTY(heat)
OP           MSG("heat.txt holds no heat number, so nothing was changed.")
OP           LET(found, -1)
OP         ELSE
OP           LOOP_ROWS dir=up from=LASTROW($ws, cHeat) to=2 var=r
OP             IF SAME(CELL($ws, r, cHeat), heat) THEN LET(found, r)
OP           IF found > 0 THEN SET_VALUE(CELL($ws, found, cRemark), "Verified")
OP       RETURN(found)
ERR      standard
FIXTURE  heat_01.xlsx sheet=Mtc cols=A:Heat No,B:MTC No,C:Supplier,D:Remark rows=9 data=A:values("H-100"x1," h-204 "x1,"H-300"x1,""x1,"H-204"x2,"H-2040"x1,"H-500"x1,"H-900"x1)+boundary("H-100"," h-204 ","H-300","","H-204","H-2040","H-500","H-204","H-900");B:seq(5001,1);C:text(10);D:values("old"x1,""x7,"Verified"x1)+boundary("old","","","","","","Verified","","")
FIXTURE  heat_01.xlsx sheet=Notes cols=A:Note rows=2 data=A:text(8)
FIXTURE  heat_01.xlsx folder=Certs files=heat.txt:"H-204",other.txt:"H-300"
FIXTURE  heat_01.xlsx folder=First files=heat.txt:"h-100"
FIXTURE  heat_01.xlsx folder=Last files=heat.txt:" H-900"
FIXTURE  heat_01.xlsx folder=NoHeat files=heat.txt:"H-777"
FIXTURE  heat_01.xlsx folder=Blank files=heat.txt:" "
FIXTURE  heat_01.xlsx folder=NoFile files=readme.txt:"no heat here"
FIXTURE  heat_01.xlsx folder=Held files=heat.txt:"H-204" locked=heat.txt
TEST     fixture=heat_01.xlsx call=VerifyHeatFromFile(Mtc, Certs) runs=2 assert=RESULT = 3 AND ERROR_COUNT = 0 AND VALUE(Mtc![Remark]3) = "Verified" AND COUNTIF(Mtc![Remark], "Verified") = 2 AND VALUE(Mtc![Remark]6) = "" AND VALUE(Mtc![Remark]9) = "" AND FILE_TEXT("Certs\heat.txt") = "H-204" AND FILE_COUNT("Certs") = PRE_FILE_COUNT("Certs")
TEST     fixture=heat_01.xlsx call=VerifyHeatFromFile(Mtc, First) assert=RESULT = 2 AND ERROR_COUNT = 0 AND VALUE(Mtc![Remark]2) = "Verified" AND COUNTIF(Mtc![Remark], "Verified") = 2 AND COUNTA(Mtc![Heat No]) = PRE_COUNTA(Mtc![Heat No])
TEST     fixture=heat_01.xlsx call=VerifyHeatFromFile(Mtc, Last) assert=RESULT = 10 AND ERROR_COUNT = 0 AND VALUE(Mtc![Remark]10) = "Verified" AND COUNTIF(Mtc![Remark], "Verified") = 2
TEST     fixture=heat_01.xlsx call=VerifyHeatFromFile(Mtc, NoHeat) assert=RESULT = 0 AND ERROR_COUNT = 0 AND COUNTIF(Mtc![Remark], "Verified") = PRE_COUNTIF(Mtc![Remark], "Verified") AND COUNTA(Mtc![Remark]) = PRE_COUNTA(Mtc![Remark])
TEST     fixture=heat_01.xlsx call=VerifyHeatFromFile(Mtc, Blank) assert=RESULT = -1 AND ERROR_COUNT = 0 AND COUNTA(Mtc![Remark]) = PRE_COUNTA(Mtc![Remark])
TEST     fixture=heat_01.xlsx call=VerifyHeatFromFile(Mtc, NoFile) assert=RESULT = -1 AND ERROR_COUNT = 0 AND COUNTA(Mtc![Remark]) = PRE_COUNTA(Mtc![Remark]) AND FILE_EXISTS("NoFile\heat.txt") = 0
TEST     fixture=heat_01.xlsx call=VerifyHeatFromFile(Mtc, Held) assert=RESULT = -1 AND ERROR_COUNT = 0 AND COUNTA(Mtc![Remark]) = PRE_COUNTA(Mtc![Remark]) AND FILE_TEXT("Held\heat.txt") = "H-204"
TEST     fixture=heat_01.xlsx call=VerifyHeatFromFile(Notes, Certs) assert=RESULT = -1 AND ERROR_COUNT = 1
