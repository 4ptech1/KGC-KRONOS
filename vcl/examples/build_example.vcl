VCL 0.1
BUILD
NAME    ReportConsolidator
TARGET  xlam
ENTRY   modMain.Run
MODULE  modMain
MODULE  modIO
FORM    frmMain CONTROLS txtFolder:TextBox btnRun:CommandButton
SIG     modIO.ListInputFiles(folderPath As String) As Variant
SIG     modMain.RunHeadless(folderPath As String) As Long
SIG     modMain.Run()
FLOW    frmMain.btnRun -> modMain.Run -> modMain.RunHeadless -> modIO.ListInputFiles
TEST    fixture=folder_01 call=modMain.RunHeadless("fixtures/folder_01") assert=COUNTA(Summary!A:A) > 1
