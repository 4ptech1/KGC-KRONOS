' =============================================================================================
' Pilot code for specs/task_01.vcl, hand-written in the shape the compiler and batch_builder.py
' will generate later: the procedure under test, then one test procedure per TEST line.
' =============================================================================================

' ---------------------------------------------------------------------------------------------
' specs/task_01.vcl: PROC BuildShortageList, KIND Function, MODULE modMaterials, RET Long
' PERF screen_off calc_manual, ERR standard. Line numbers are there so Erl can report them.
' ---------------------------------------------------------------------------------------------
Public Function BuildShortageList(ByVal bom As Worksheet, ByVal grn As Worksheet, ByVal out As Worksheet) As Long
    Dim received As Object
    Dim bomPart As Long
    Dim bomReq As Long
    Dim grnPart As Long
    Dim grnQty As Long
    Dim partNo As String
    Dim got As Double
    Dim outRow As Long
    Dim r As Long
    Dim preFailed As String
    Dim perfOn As Boolean
    Dim prevScreen As Boolean
    Dim prevCalc As Long
    On Error GoTo Fail
    ' PRE: checked before any change; a failure logs once, changes nothing and returns -1
10  If ForgeCol(bom, "Part No") = 0 Then preFailed = "COL($bom, ""Part No"") > 0"
20  If Len(preFailed) = 0 And ForgeCol(bom, "Required Qty") = 0 Then preFailed = "COL($bom, ""Required Qty"") > 0"
30  If Len(preFailed) = 0 And ForgeCol(grn, "Part No") = 0 Then preFailed = "COL($grn, ""Part No"") > 0"
40  If Len(preFailed) = 0 And ForgeCol(grn, "Received Qty") = 0 Then preFailed = "COL($grn, ""Received Qty"") > 0"
50  If Len(preFailed) > 0 Then
60      ForgeLog "BuildShortageList", 0, 0, "PRE failed: " & preFailed
70      BuildShortageList = -1
80      Exit Function
    End If
    ' PERF screen_off calc_manual
90  prevScreen = Application.ScreenUpdating
100 prevCalc = Application.Calculation
110 perfOn = True
120 Application.ScreenUpdating = False
130 Application.Calculation = xlCalculationManual
    ' VAR received As Dictionary (keys ignore case)
140 Set received = CreateObject("Scripting.Dictionary")
150 received.CompareMode = vbTextCompare
160 bomPart = ForgeCol(bom, "Part No")
170 bomReq = ForgeCol(bom, "Required Qty")
180 grnPart = ForgeCol(grn, "Part No")
190 grnQty = ForgeCol(grn, "Received Qty")
200 out.UsedRange.Clear
210 out.Cells(1, 1).Value = "Part No"
220 out.Cells(1, 2).Value = "Required"
230 out.Cells(1, 3).Value = "Received"
240 out.Cells(1, 4).Value = "Shortage"
250 out.Cells(1, 5).Value = "Remark"
260 For r = 2 To ForgeLastRow(grn, grnPart)
270     If Not ForgeIsEmpty(grn.Cells(r, grnPart).Value) And ForgeIsNumber(grn.Cells(r, grnQty).Value) Then
280         ForgeDictAdd received, Trim$(CStr(grn.Cells(r, grnPart).Value)), grn.Cells(r, grnQty).Value, "sum"
        End If
    Next r
290 outRow = 2
300 For r = 2 To ForgeLastRow(bom, bomPart)
310     If Not ForgeIsEmpty(bom.Cells(r, bomPart).Value) Then
320         partNo = Trim$(CStr(bom.Cells(r, bomPart).Value))
330         got = 0
340         If received.Exists(partNo) Then got = received.Item(partNo)
350         If Not ForgeIsNumber(bom.Cells(r, bomReq).Value) Then
360             out.Cells(outRow, 1).Value = partNo
370             out.Cells(outRow, 3).Value = got
380             out.Cells(outRow, 5).Value = "Required qty missing"
390             outRow = outRow + 1
            Else
400             If got < CDbl(bom.Cells(r, bomReq).Value) Then
410                 out.Cells(outRow, 1).Value = partNo
420                 out.Cells(outRow, 2).Value = bom.Cells(r, bomReq).Value
430                 out.Cells(outRow, 3).Value = got
440                 out.Cells(outRow, 4).Value = CDbl(bom.Cells(r, bomReq).Value) - got
450                 If got = 0 Then out.Cells(outRow, 5).Value = "Not received"
460                 If got > 0 Then out.Cells(outRow, 5).Value = "Short"
470                 outRow = outRow + 1
                End If
            End If
        End If
    Next r
480 BuildShortageList = outRow - 2
Done:
    On Error Resume Next
    If perfOn Then
        Application.ScreenUpdating = prevScreen
        Application.Calculation = prevCalc
    End If
    Exit Function
Fail:
    ForgeLog "BuildShortageList", Err.Number, Erl, Err.Description
    Resume Done
End Function

' ---------------------------------------------------------------------------------------------
' TEST B000-T01, specs/task_01.vcl line 60
' fixture=bom_grn_01.xlsx call=BuildShortageList(BOM, GRN, Shortages) runs=2
' assert=RESULT >= 5 AND COUNTA(Shortages![Part No]) = RESULT + 1 AND
'        COUNTIF(Shortages![Shortage], "<0") = 0 AND ERROR_COUNT = 0
' ---------------------------------------------------------------------------------------------
Private Sub ForgeTest_B000_T01()
    Dim wb As Workbook
    On Error GoTo Fail
10  Set wb = ForgeBeginTest("B000-T01", "task_01")
20  ForgeFixture_bom_grn_01 wb
30  ForgeCall "BuildShortageList", True, 2, _
        Array(wb.Worksheets("BOM"), wb.Worksheets("GRN"), wb.Worksheets("Shortages"))
40  ForgeCheck "RESULT >= 5", ForgeResult(), ">=", 5
50  ForgeCheck "COUNTA(Shortages![Part No]) = RESULT + 1", _
        ForgeMeasure("COUNTA(Shortages![Part No])", ForgeCountA(ForgeHeaderRange(wb, "Shortages", "Part No"))), _
        "=", ForgeResult() + 1
60  ForgeCheck "COUNTIF(Shortages![Shortage], ""<0"") = 0", _
        ForgeMeasure("COUNTIF(Shortages![Shortage], ""<0"")", _
                     ForgeCountIf(ForgeHeaderRange(wb, "Shortages", "Shortage"), "<0")), _
        "=", 0
70  ForgeCheck "ERROR_COUNT = 0", ForgeErrorCount(), "=", 0
80  ForgeEndTest
    Exit Sub
Fail:
    ForgeTestError Err.Number, Erl, Err.Description
End Sub

' ---------------------------------------------------------------------------------------------
' TEST B000-T02, specs/task_01.vcl line 61
' fixture=bom_grn_01.xlsx call=BuildShortageList(BOM, GRN, Shortages)
' assert=COUNTIF(Shortages![Remark], "Required qty missing") = 3
' ---------------------------------------------------------------------------------------------
Private Sub ForgeTest_B000_T02()
    Dim wb As Workbook
    On Error GoTo Fail
10  Set wb = ForgeBeginTest("B000-T02", "task_01")
20  ForgeFixture_bom_grn_01 wb
30  ForgeCall "BuildShortageList", True, 1, _
        Array(wb.Worksheets("BOM"), wb.Worksheets("GRN"), wb.Worksheets("Shortages"))
40  ForgeCheck "COUNTIF(Shortages![Remark], ""Required qty missing"") = 3", _
        ForgeMeasure("COUNTIF(Shortages![Remark], ""Required qty missing"")", _
                     ForgeCountIf(ForgeHeaderRange(wb, "Shortages", "Remark"), "Required qty missing")), _
        "=", 3
50  ForgeEndTest
    Exit Sub
Fail:
    ForgeTestError Err.Number, Erl, Err.Description
End Sub

' ---------------------------------------------------------------------------------------------
' TEST B000-T03, specs/task_01.vcl line 62
' fixture=bom_grn_01.xlsx call=BuildShortageList(BOM, GRN, Shortages)
' assert=SUM(BOM![Required Qty]) = PRE_SUM(BOM![Required Qty]) AND
'        SUM(GRN![Received Qty]) = PRE_SUM(GRN![Received Qty])
' ---------------------------------------------------------------------------------------------
Private Sub ForgeTest_B000_T03()
    Dim wb As Workbook
    Dim pre1 As Variant
    Dim pre2 As Variant
    On Error GoTo Fail
10  Set wb = ForgeBeginTest("B000-T03", "task_01")
20  ForgeFixture_bom_grn_01 wb
30  pre1 = ForgeMeasure("PRE_SUM(BOM![Required Qty])", ForgeSum(ForgeHeaderRange(wb, "BOM", "Required Qty")))
40  pre2 = ForgeMeasure("PRE_SUM(GRN![Received Qty])", ForgeSum(ForgeHeaderRange(wb, "GRN", "Received Qty")))
50  ForgeCall "BuildShortageList", True, 1, _
        Array(wb.Worksheets("BOM"), wb.Worksheets("GRN"), wb.Worksheets("Shortages"))
60  ForgeCheck "SUM(BOM![Required Qty]) = PRE_SUM(BOM![Required Qty])", _
        ForgeMeasure("SUM(BOM![Required Qty])", ForgeSum(ForgeHeaderRange(wb, "BOM", "Required Qty"))), _
        "=", pre1
70  ForgeCheck "SUM(GRN![Received Qty]) = PRE_SUM(GRN![Received Qty])", _
        ForgeMeasure("SUM(GRN![Received Qty])", ForgeSum(ForgeHeaderRange(wb, "GRN", "Received Qty"))), _
        "=", pre2
80  ForgeEndTest
    Exit Sub
Fail:
    ForgeTestError Err.Number, Erl, Err.Description
End Sub

' ---------------------------------------------------------------------------------------------
' TEST B000-T04, specs/task_01.vcl line 63 (wrong sheet: PRE must fail)
' fixture=bom_grn_01.xlsx call=BuildShortageList(Shortages, GRN, Shortages)
' assert=RESULT = -1 AND ERROR_COUNT = 1 AND
'        COUNTA(Shortages![Part No]) = PRE_COUNTA(Shortages![Part No])
' ---------------------------------------------------------------------------------------------
Private Sub ForgeTest_B000_T04()
    Dim wb As Workbook
    Dim pre1 As Variant
    On Error GoTo Fail
10  Set wb = ForgeBeginTest("B000-T04", "task_01")
20  ForgeFixture_bom_grn_01 wb
30  pre1 = ForgeMeasure("PRE_COUNTA(Shortages![Part No])", ForgeCountA(ForgeHeaderRange(wb, "Shortages", "Part No")))
40  ForgeCall "BuildShortageList", True, 1, _
        Array(wb.Worksheets("Shortages"), wb.Worksheets("GRN"), wb.Worksheets("Shortages"))
50  ForgeCheck "RESULT = -1", ForgeResult(), "=", -1
60  ForgeCheck "ERROR_COUNT = 1", ForgeErrorCount(), "=", 1
70  ForgeCheck "COUNTA(Shortages![Part No]) = PRE_COUNTA(Shortages![Part No])", _
        ForgeMeasure("COUNTA(Shortages![Part No])", ForgeCountA(ForgeHeaderRange(wb, "Shortages", "Part No"))), _
        "=", pre1
80  ForgeEndTest
    Exit Sub
Fail:
    ForgeTestError Err.Number, Erl, Err.Description
End Sub
