Option Explicit

' =============================================================================================
' VBA Forge batch runner engine (harness/ForgeTest.bas)
'
' This file is not pasted on its own. harness/assemble.py appends a batch to it (the test list,
' the fixture builders and the procedures under test) and writes one paste-ready module to
' batches/. The batch must define ForgeBatchName, ForgeBatchTestCount and ForgeBatchTests.
'
' RunBatch (Alt+F8) runs every test of the batch. Each test gets its own temporary workbook:
' the fixtures are written into it, PRE_ measures are taken, the procedure is called with
' Application.Run (runs=n times), the assertions are checked, and the workbook is closed
' without saving. Results go to the Results sheet of this workbook, ForgeLog entries to the
' ForgeLog sheet. One failing or erroring test never stops the others.
'
' ForgeReset (Alt+F8) puts Excel's settings back and closes leftover test workbooks after a
' batch was stopped with Ctrl+Break and End.
'
' Excel 2016, no References: Scripting.Dictionary is late-bound. No Select or Activate.
' =============================================================================================

Private Const FORGE_RESULTS As String = "Results"
Private Const FORGE_LOG_SHEET As String = "ForgeLog"
Private Const FORGE_TEMP_NAME As String = "ForgeTempWorkbook"
Private Const FORGE_BLANK As String = "ForgeBlank"

Private mResults As Worksheet       ' Results sheet in this workbook
Private mLogSheet As Worksheet      ' ForgeLog sheet in this workbook
Private mRow As Long                ' Results row of the current test
Private mTestNo As Long             ' number of the current test in this batch
Private mTestId As String
Private mWb As Workbook             ' temporary workbook of the current test
Private mStart As Double            ' Timer when the current test started
Private mResult As Variant          ' RESULT: the Function's return value from the last run
Private mLogCount As Long           ' ForgeLog entries during the current test (ERROR_COUNT)
Private mLogNo As Long              ' last ForgeLog entry of the current test
Private mLogLine As Long
Private mLogText As String
Private mFailed As String           ' first assertion of the current test that failed
Private mDetail As String           ' measures of the current test, for the Description column
Private mPassCount As Long
Private mFailCount As Long
Private mErrorCount As Long
Private mSaved As Boolean           ' Excel settings saved by RunBatch, restored at the end
Private mSavedScreen As Boolean
Private mSavedCalc As Long
Private mSavedEvents As Boolean
Private mSavedAlerts As Boolean

' ---------------------------------------------------------------------------------------------
' Entry points (the only two macros listed by Alt+F8)
' ---------------------------------------------------------------------------------------------

Public Sub RunBatch()
    Dim summary As String
    Dim errText As String
    On Error GoTo Fail
    ForgeSaveAppState
    ForgeCloseLeftovers
    ForgePrepareSheets
    mTestNo = 0
    mPassCount = 0
    mFailCount = 0
    mErrorCount = 0
    ForgeApplyRunnerState
    ForgeBatchTests
    ForgeRestoreAppState
    ForgeFinishSheets
    summary = "VBA Forge batch " & ForgeBatchName() & " finished." & vbCrLf & vbCrLf
    summary = summary & mTestNo & " tests: " & mPassCount & " PASS, " & mFailCount & " FAIL, "
    summary = summary & mErrorCount & " ERROR." & vbCrLf & vbCrLf
    summary = summary & "Open the Results sheet and copy it back to Claude."
    MsgBox summary, vbInformation, "VBA Forge"
    Exit Sub
Fail:
    errText = "error " & Err.Number & ": " & Err.Description
    ForgeRestoreAppState
    ForgeCloseTemp
    MsgBox "VBA Forge stopped early (" & errText & ")." & vbCrLf & _
           "Send Claude this message and the Results sheet.", vbExclamation, "VBA Forge"
End Sub

Public Sub ForgeReset()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    ForgeCloseLeftovers
    MsgBox "VBA Forge: Excel settings are back to normal and leftover test workbooks are closed.", _
           vbInformation, "VBA Forge"
End Sub

' ---------------------------------------------------------------------------------------------
' Excel settings
' ---------------------------------------------------------------------------------------------

Private Sub ForgeSaveAppState()
    mSavedScreen = Application.ScreenUpdating
    mSavedCalc = Application.Calculation
    mSavedEvents = Application.EnableEvents
    mSavedAlerts = Application.DisplayAlerts
    mSaved = True
End Sub

' Settings while tests run: no screen redraw, no alerts, the user's calculation mode and events.
Private Sub ForgeApplyRunnerState()
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    If mSaved Then
        Application.EnableEvents = mSavedEvents
        If Application.Calculation <> mSavedCalc Then Application.Calculation = mSavedCalc
    End If
End Sub

Private Sub ForgeRestoreAppState()
    On Error Resume Next
    If mSaved Then
        Application.ScreenUpdating = mSavedScreen
        Application.EnableEvents = mSavedEvents
        Application.DisplayAlerts = mSavedAlerts
        If Application.Calculation <> mSavedCalc Then Application.Calculation = mSavedCalc
    Else
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
    End If
    Application.StatusBar = False
End Sub

' ---------------------------------------------------------------------------------------------
' Results and ForgeLog sheets
' ---------------------------------------------------------------------------------------------

Private Sub ForgePrepareSheets()
    Set mResults = ForgeSheetIn(ThisWorkbook, FORGE_RESULTS)
    mResults.Cells.Clear
    mResults.Range("A1:G1").Value = Array("Test ID", "Spec", "Pass/Fail", "Error No", "Line", "Description", "Seconds")
    mResults.Range("A1:G1").Font.Bold = True
    Set mLogSheet = ForgeSheetIn(ThisWorkbook, FORGE_LOG_SHEET)
    mLogSheet.Cells.Clear
    mLogSheet.Range("A1:F1").Value = Array("Time", "Test ID", "Procedure", "Error No", "Line", "Message")
    mLogSheet.Range("A1:F1").Font.Bold = True
    mRow = 1
End Sub

Private Sub ForgeFinishSheets()
    On Error Resume Next
    mResults.Columns("A:E").AutoFit
    mResults.Columns("F").ColumnWidth = 100
    mResults.Columns("G").AutoFit
    mLogSheet.Columns("A").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    mLogSheet.Columns("A:F").AutoFit
End Sub

Private Function ForgeSheetIn(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            Set ForgeSheetIn = ws
            Exit Function
        End If
    Next ws
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheetName
    Set ForgeSheetIn = ws
End Function

Private Function ForgeHasSheet(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            ForgeHasSheet = True
            Exit Function
        End If
    Next ws
End Function

Private Sub ForgeWriteResult(ByVal outcome As String, ByVal errNo As Long, ByVal lineNo As Long, _
                             ByVal descText As String)
    Dim secs As Double
    secs = Timer - mStart
    If secs < 0 Then secs = secs + 86400
    mResults.Cells(mRow, 3).Value = outcome
    If errNo <> 0 Then mResults.Cells(mRow, 4).Value = errNo
    If lineNo <> 0 Then mResults.Cells(mRow, 5).Value = lineNo
    mResults.Cells(mRow, 6).Value = "'" & ForgeOneLine(descText)
    mResults.Cells(mRow, 7).Value = Round(secs, 2)
    Select Case outcome
        Case "PASS"
            mPassCount = mPassCount + 1
        Case "FAIL"
            mFailCount = mFailCount + 1
        Case Else
            mErrorCount = mErrorCount + 1
    End Select
End Sub

' ---------------------------------------------------------------------------------------------
' Test lifecycle (called by the generated test procedures)
' ---------------------------------------------------------------------------------------------

Private Function ForgeBeginTest(ByVal testId As String, ByVal specName As String) As Workbook
    mTestNo = mTestNo + 1
    mTestId = testId
    mRow = mRow + 1
    mResult = Empty
    mLogCount = 0
    mLogNo = 0
    mLogLine = 0
    mLogText = ""
    mFailed = ""
    mDetail = ""
    mStart = Timer
    ' The test id goes in first, so a hang shows which test it was.
    mResults.Cells(mRow, 1).Value = "'" & testId
    mResults.Cells(mRow, 2).Value = "'" & specName
    mResults.Cells(mRow, 3).Value = "RUNNING"
    Application.StatusBar = "VBA Forge " & ForgeBatchName() & ": running " & testId & " (" & mTestNo & _
                            " of " & ForgeBatchTestCount() & ")"
    Set mWb = Workbooks.Add(xlWBATWorksheet)
    mWb.Worksheets(1).Name = FORGE_BLANK
    mWb.Names.Add Name:=FORGE_TEMP_NAME, RefersTo:="=TRUE"
    Set ForgeBeginTest = mWb
End Function

Private Sub ForgeEndTest()
    Dim outcome As String
    Dim descText As String
    If Len(mFailed) = 0 Then
        outcome = "PASS"
        descText = ForgeSummary()
    Else
        outcome = "FAIL"
        descText = "failed: " & mFailed & " | " & ForgeSummary()
    End If
    ForgeWriteResult outcome, mLogNo, mLogLine, descText
    ForgeCloseTemp
    ForgeApplyRunnerState
End Sub

' Called from a test procedure's error handler: record ERROR and move on to the next test.
Private Sub ForgeTestError(ByVal errNo As Long, ByVal lineNo As Long, ByVal message As String)
    On Error Resume Next
    ForgeWriteResult "ERROR", errNo, lineNo, "error in test step " & lineNo & ": " & message & " | " & ForgeSummary()
    ForgeCloseTemp
    ForgeApplyRunnerState
End Sub

Private Function ForgeSummary() As String
    Dim s As String
    s = "RESULT=" & ForgeShow(mResult) & "; ERROR_COUNT=" & mLogCount & mDetail
    If Len(mLogText) > 0 Then s = s & " | last ForgeLog: " & mLogText
    ForgeSummary = s
End Function

Private Sub ForgeCloseTemp()
    On Error Resume Next
    If Not mWb Is Nothing Then mWb.Close SaveChanges:=False
    Set mWb = Nothing
End Sub

' Closes test workbooks left open by an interrupted run (they carry the ForgeTempWorkbook name).
Private Sub ForgeCloseLeftovers()
    Dim i As Long
    On Error Resume Next
    For i = Workbooks.Count To 1 Step -1
        If ForgeIsTempWorkbook(Workbooks(i)) Then Workbooks(i).Close SaveChanges:=False
    Next i
End Sub

Private Function ForgeIsTempWorkbook(ByVal wb As Workbook) As Boolean
    Dim nm As Excel.Name
    On Error Resume Next
    For Each nm In wb.Names
        If nm.Name = FORGE_TEMP_NAME Then
            ForgeIsTempWorkbook = True
            Exit Function
        End If
    Next nm
End Function

' Calls the procedure under test runs times, with Application.Run. RESULT is the last return value.
Private Sub ForgeCall(ByVal procName As String, ByVal isFunction As Boolean, ByVal runs As Long, _
                      ByVal args As Variant)
    Dim i As Long
    Dim macroName As String
    Dim calcBefore As Long
    Dim eventsBefore As Boolean
    Dim returned As Variant
    macroName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & procName
    For i = 1 To runs
        calcBefore = Application.Calculation
        eventsBefore = Application.EnableEvents
        returned = ForgeRunMacro(macroName, isFunction, args)
        If isFunction Then mResult = returned
        If Application.Calculation <> calcBefore Then ForgeFail "(runner) run " & i & " left Calculation changed"
        If Application.EnableEvents <> eventsBefore Then ForgeFail "(runner) run " & i & " left EnableEvents changed"
    Next i
End Sub

Private Function ForgeRunMacro(ByVal macroName As String, ByVal isFunction As Boolean, _
                               ByVal args As Variant) As Variant
    Dim b As Long
    Dim n As Long
    b = LBound(args)
    n = UBound(args) - b + 1
    If isFunction Then
        Select Case n
            Case 0
                ForgeRunMacro = Application.Run(macroName)
            Case 1
                ForgeRunMacro = Application.Run(macroName, args(b))
            Case 2
                ForgeRunMacro = Application.Run(macroName, args(b), args(b + 1))
            Case 3
                ForgeRunMacro = Application.Run(macroName, args(b), args(b + 1), args(b + 2))
            Case 4
                ForgeRunMacro = Application.Run(macroName, args(b), args(b + 1), args(b + 2), args(b + 3))
            Case 5
                ForgeRunMacro = Application.Run(macroName, args(b), args(b + 1), args(b + 2), args(b + 3), _
                                                args(b + 4))
            Case 6
                ForgeRunMacro = Application.Run(macroName, args(b), args(b + 1), args(b + 2), args(b + 3), _
                                                args(b + 4), args(b + 5))
            Case Else
                Err.Raise vbObjectError + 513, "ForgeRunMacro", "the runner passes at most 6 arguments"
        End Select
    Else
        Select Case n
            Case 0
                Application.Run macroName
            Case 1
                Application.Run macroName, args(b)
            Case 2
                Application.Run macroName, args(b), args(b + 1)
            Case 3
                Application.Run macroName, args(b), args(b + 1), args(b + 2)
            Case 4
                Application.Run macroName, args(b), args(b + 1), args(b + 2), args(b + 3)
            Case 5
                Application.Run macroName, args(b), args(b + 1), args(b + 2), args(b + 3), args(b + 4)
            Case 6
                Application.Run macroName, args(b), args(b + 1), args(b + 2), args(b + 3), args(b + 4), _
                                args(b + 5)
            Case Else
                Err.Raise vbObjectError + 513, "ForgeRunMacro", "the runner passes at most 6 arguments"
        End Select
    End If
End Function

' ---------------------------------------------------------------------------------------------
' Assertions
' ---------------------------------------------------------------------------------------------

Private Function ForgeResult() As Variant
    ForgeResult = mResult
End Function

Private Function ForgeErrorCount() As Long
    ForgeErrorCount = mLogCount
End Function

' Records a measure (once per name) for the Description column and passes the value through.
Private Function ForgeMeasure(ByVal measureName As String, ByVal measured As Variant) As Variant
    If InStr(1, mDetail, "; " & measureName & "=", vbBinaryCompare) = 0 Then
        mDetail = mDetail & "; " & measureName & "=" & ForgeShow(measured)
    End If
    ForgeMeasure = measured
End Function

Private Sub ForgeCheck(ByVal checkName As String, ByVal leftValue As Variant, ByVal op As String, _
                       ByVal rightValue As Variant)
    If Not ForgeCompare(leftValue, op, rightValue) Then
        ForgeFail checkName & " [" & ForgeShow(leftValue) & " " & op & " " & ForgeShow(rightValue) & "]"
    End If
End Sub

Private Sub ForgeFail(ByVal reason As String)
    If Len(mFailed) = 0 Then mFailed = reason
End Sub

Private Function ForgeCompare(ByVal leftValue As Variant, ByVal op As String, ByVal rightValue As Variant) As Boolean
    Dim cmp As Long
    Dim x As Double
    Dim y As Double
    If IsObject(leftValue) Or IsObject(rightValue) Then Exit Function
    If IsError(leftValue) Or IsError(rightValue) Then Exit Function
    If IsNull(leftValue) Or IsNull(rightValue) Then Exit Function
    If ForgeIsNumericValue(leftValue) And ForgeIsNumericValue(rightValue) Then
        x = CDbl(leftValue)
        y = CDbl(rightValue)
        If Abs(x - y) <= 0.000000001 * (1 + Abs(x) + Abs(y)) Then
            cmp = 0
        ElseIf x < y Then
            cmp = -1
        Else
            cmp = 1
        End If
    Else
        cmp = StrComp(CStr(leftValue), CStr(rightValue), vbTextCompare)
    End If
    Select Case op
        Case "="
            ForgeCompare = (cmp = 0)
        Case "<>"
            ForgeCompare = (cmp <> 0)
        Case "<"
            ForgeCompare = (cmp < 0)
        Case "<="
            ForgeCompare = (cmp <= 0)
        Case ">"
            ForgeCompare = (cmp > 0)
        Case ">="
            ForgeCompare = (cmp >= 0)
        Case Else
            Err.Raise vbObjectError + 514, "ForgeCompare", "unknown operator " & op
    End Select
End Function

Private Function ForgeIsNumericValue(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbEmpty, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte, vbDate
            ForgeIsNumericValue = True
    End Select
End Function

Private Function ForgeShow(ByVal v As Variant) As String
    If IsObject(v) Then
        ForgeShow = "(object)"
    ElseIf IsError(v) Then
        ForgeShow = "(" & CStr(v) & ")"
    ElseIf IsNull(v) Then
        ForgeShow = "(null)"
    ElseIf IsEmpty(v) Then
        ForgeShow = "(empty)"
    ElseIf VarType(v) = vbString Then
        ForgeShow = """" & v & """"
    ElseIf VarType(v) = vbDate Then
        ForgeShow = Format$(v, "yyyy-mm-dd")
    ElseIf VarType(v) = vbBoolean Then
        If v Then ForgeShow = "TRUE" Else ForgeShow = "FALSE"
    Else
        ForgeShow = Trim$(Str$(v))
    End If
End Function

Private Function ForgeOneLine(ByVal s As String) As String
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    If Len(s) > 2000 Then s = Left$(s, 2000) & "..."
    ForgeOneLine = s
End Function

' ---------------------------------------------------------------------------------------------
' Assertion measures. <Sheet>![Header] is the whole column under that row-1 header, header row
' included (so COUNTA counts the header, which the specs' "= RESULT + 1" checks rely on).
' ---------------------------------------------------------------------------------------------

Private Function ForgeHeaderRange(ByVal wb As Workbook, ByVal sheetName As String, _
                                  ByVal headerText As String) As Range
    Dim ws As Worksheet
    Dim col As Long
    Set ws = wb.Worksheets(sheetName)
    col = ForgeCol(ws, headerText)
    If col = 0 Then
        Err.Raise vbObjectError + 515, "ForgeHeaderRange", "sheet " & sheetName & " has no [" & headerText & "] header"
    End If
    Set ForgeHeaderRange = ws.Columns(col)
End Function

Private Function ForgeHeaderCell(ByVal wb As Workbook, ByVal sheetName As String, ByVal headerText As String, _
                                 ByVal rowNo As Long) As Range
    Set ForgeHeaderCell = ForgeHeaderRange(wb, sheetName, headerText).Cells(rowNo, 1)
End Function

Private Function ForgeSheetRange(ByVal wb As Workbook, ByVal sheetName As String, ByVal refText As String) As Range
    Set ForgeSheetRange = wb.Worksheets(sheetName).Range(refText)
End Function

Private Function ForgeCountA(ByVal rng As Range) As Double
    ForgeCountA = Application.WorksheetFunction.CountA(rng)
End Function

Private Function ForgeCountIf(ByVal rng As Range, ByVal criteria As Variant) As Double
    ForgeCountIf = Application.WorksheetFunction.CountIf(rng, criteria)
End Function

Private Function ForgeSum(ByVal rng As Range) As Double
    ForgeSum = Application.WorksheetFunction.Sum(rng)
End Function

Private Function ForgeCountIfs(ParamArray parts() As Variant) As Double
    Dim wf As WorksheetFunction
    Set wf = Application.WorksheetFunction
    Select Case UBound(parts) + 1
        Case 2
            ForgeCountIfs = wf.CountIfs(parts(0), parts(1))
        Case 4
            ForgeCountIfs = wf.CountIfs(parts(0), parts(1), parts(2), parts(3))
        Case 6
            ForgeCountIfs = wf.CountIfs(parts(0), parts(1), parts(2), parts(3), parts(4), parts(5))
        Case 8
            ForgeCountIfs = wf.CountIfs(parts(0), parts(1), parts(2), parts(3), parts(4), parts(5), _
                                        parts(6), parts(7))
        Case Else
            Err.Raise vbObjectError + 516, "ForgeCountIfs", "COUNTIFS takes 1 to 4 range, criteria pairs"
    End Select
End Function

Private Function ForgeSumIfs(ParamArray parts() As Variant) As Double
    Dim wf As WorksheetFunction
    Set wf = Application.WorksheetFunction
    Select Case UBound(parts) + 1
        Case 3
            ForgeSumIfs = wf.SumIfs(parts(0), parts(1), parts(2))
        Case 5
            ForgeSumIfs = wf.SumIfs(parts(0), parts(1), parts(2), parts(3), parts(4))
        Case 7
            ForgeSumIfs = wf.SumIfs(parts(0), parts(1), parts(2), parts(3), parts(4), parts(5), parts(6))
        Case 9
            ForgeSumIfs = wf.SumIfs(parts(0), parts(1), parts(2), parts(3), parts(4), parts(5), parts(6), _
                                    parts(7), parts(8))
        Case Else
            Err.Raise vbObjectError + 517, "ForgeSumIfs", "SUMIFS takes a sum range and 1 to 4 range, criteria pairs"
    End Select
End Function

Private Function ForgeValue(ByVal rng As Range) As Variant
    ForgeValue = rng.Cells(1, 1).Value
End Function

Private Function ForgeRowsUsed(ByVal wb As Workbook, ByVal sheetName As String) As Long
    Dim found As Range
    Set found = wb.Worksheets(sheetName).Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                                    SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not found Is Nothing Then ForgeRowsUsed = found.Row
End Function

Private Function ForgeSheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    ForgeSheetExists = ForgeHasSheet(wb, sheetName)
End Function

Private Function ForgeFileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    ForgeFileExists = (Len(Dir$(filePath)) > 0)
End Function

' Cells in rng (within the used range) whose fill is the #RRGGBB colour.
Private Function ForgeFillCount(ByVal rng As Range, ByVal hexColor As String) As Long
    Dim area As Range
    Dim cell As Range
    Dim target As Long
    Dim n As Long
    target = ForgeHexToColor(hexColor)
    Set area = Application.Intersect(rng, rng.Worksheet.UsedRange)
    If area Is Nothing Then Exit Function
    For Each cell In area.Cells
        If cell.Interior.Pattern <> xlPatternNone Then
            If cell.Interior.Color = target Then n = n + 1
        End If
    Next cell
    ForgeFillCount = n
End Function

Private Function ForgeHexToColor(ByVal hexColor As String) As Long
    ForgeHexToColor = RGB(CLng("&H" & Mid$(hexColor, 2, 2)), CLng("&H" & Mid$(hexColor, 4, 2)), _
                          CLng("&H" & Mid$(hexColor, 6, 2)))
End Function

' ---------------------------------------------------------------------------------------------
' Fixtures (the builders are generated by harness/fixtures.py)
' ---------------------------------------------------------------------------------------------

' The first fixture sheet reuses the new workbook's only sheet; later ones are added at the end.
Private Function ForgeFixtureSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    If ForgeHasSheet(wb, FORGE_BLANK) Then
        Set ws = wb.Worksheets(FORGE_BLANK)
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    End If
    ws.Name = sheetName
    Set ForgeFixtureSheet = ws
End Function

' Writes one row of literal values. fixtures.py already put a ' in front of text that Excel would
' convert ("25" stays text, "1-2" doesn't turn into a date). Empty values leave the cell blank.
Private Sub ForgeWriteRow(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal rowValues As Variant)
    Dim i As Long
    Dim n As Long
    Dim rowData() As Variant
    n = UBound(rowValues) - LBound(rowValues) + 1
    ReDim rowData(1 To 1, 1 To n)
    For i = 1 To n
        rowData(1, i) = rowValues(LBound(rowValues) + i - 1)
    Next i
    ws.Range(ws.Cells(rowNo, 1), ws.Cells(rowNo, n)).Value = rowData
End Sub

' ---------------------------------------------------------------------------------------------
' Helpers for generated procedures: VCL ops and value functions
' ---------------------------------------------------------------------------------------------

' The standard error handler and failed PREs write here. ERROR_COUNT counts these entries.
Public Sub ForgeLog(ByVal procName As String, ByVal errNo As Long, ByVal lineNo As Long, ByVal message As String)
    Dim r As Long
    mLogCount = mLogCount + 1
    mLogNo = errNo
    mLogLine = lineNo
    mLogText = procName & ": " & message
    On Error Resume Next
    If mLogSheet Is Nothing Then
        Debug.Print "ForgeLog: " & mLogText
        Exit Sub
    End If
    r = mLogSheet.Cells(mLogSheet.Rows.Count, 1).End(xlUp).Row + 1
    mLogSheet.Cells(r, 1).Value = Now
    mLogSheet.Cells(r, 2).Value = "'" & mTestId
    mLogSheet.Cells(r, 3).Value = "'" & procName
    mLogSheet.Cells(r, 4).Value = errNo
    mLogSheet.Cells(r, 5).Value = lineNo
    mLogSheet.Cells(r, 6).Value = "'" & ForgeOneLine(message)
End Sub

' COL: column number of a row-1 header, ignoring case and extra spaces; 0 if it is missing.
Public Function ForgeCol(ByVal ws As Worksheet, ByVal headerText As String) As Long
    Dim c As Long
    Dim lastCol As Long
    Dim want As String
    Dim v As Variant
    want = ForgeNormHeader(headerText)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        v = ws.Cells(1, c).Value
        If Not IsError(v) Then
            If ForgeNormHeader(CStr(v)) = want Then
                ForgeCol = c
                Exit Function
            End If
        End If
    Next c
End Function

Private Function ForgeNormHeader(ByVal s As String) As String
    Dim i As Long
    s = Trim$(s)
    For i = 1 To Len(s)
        If InStr(s, "  ") = 0 Then Exit For
        s = Replace(s, "  ", " ")
    Next i
    ForgeNormHeader = LCase$(s)
End Function

' LASTROW / LASTCOL
Public Function ForgeLastRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    ForgeLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

Public Function ForgeLastCol(ByVal ws As Worksheet, ByVal rowNo As Long) As Long
    ForgeLastCol = ws.Cells(rowNo, ws.Columns.Count).End(xlToLeft).Column
End Function

' ROW_RANGE
Public Function ForgeRowRange(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal fromCol As Long, _
                              ByVal toCol As Long) As Range
    Set ForgeRowRange = ws.Range(ws.Cells(rowNo, fromCol), ws.Cells(rowNo, toCol))
End Function

' COPY_ROW: values of columns 1 to cols
Public Sub ForgeCopyRow(ByVal fromWs As Worksheet, ByVal fromRow As Long, ByVal toWs As Worksheet, _
                        ByVal toRow As Long, ByVal cols As Long)
    toWs.Range(toWs.Cells(toRow, 1), toWs.Cells(toRow, cols)).Value = _
        fromWs.Range(fromWs.Cells(fromRow, 1), fromWs.Cells(fromRow, cols)).Value
End Sub

' IS_EMPTY: empty cells, empty text and text made only of spaces
Public Function ForgeIsEmpty(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then
        ForgeIsEmpty = True
    ElseIf VarType(v) = vbString Then
        ForgeIsEmpty = (Len(Trim$(v)) = 0)
    End If
End Function

' IS_NUMBER: numbers, and text that is a plain number ("25"); false for blanks and dates
Public Function ForgeIsNumber(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            ForgeIsNumber = True
        Case vbString
            ForgeIsNumber = ForgeIsPlainNumber(Trim$(v))
    End Select
End Function

' IS_DATE: only real Excel date values
Public Function ForgeIsDate(ByVal v As Variant) As Boolean
    ForgeIsDate = (VarType(v) = vbDate)
End Function

Private Function ForgeIsPlainNumber(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    Dim digits As Long
    Dim dots As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits + 1
        ElseIf ch = "." Then
            dots = dots + 1
            If dots > 1 Then Exit Function
        ElseIf (ch = "-" Or ch = "+") And i = 1 Then
            ' a leading sign is fine
        Else
            Exit Function
        End If
    Next i
    ForgeIsPlainNumber = (digits > 0)
End Function

' DICT_ADD with mode overwrite | keep | sum | count (dict is a late-bound Scripting.Dictionary)
Public Sub ForgeDictAdd(ByVal dict As Object, ByVal dictKey As Variant, ByVal itemValue As Variant, _
                        ByVal addMode As String)
    If Not dict.Exists(dictKey) Then
        Select Case addMode
            Case "count"
                dict.Add dictKey, 1
            Case "sum"
                dict.Add dictKey, CDbl(itemValue)
            Case "overwrite", "keep"
                dict.Add dictKey, itemValue
            Case Else
                Err.Raise vbObjectError + 518, "ForgeDictAdd", "unknown DICT_ADD mode " & addMode
        End Select
        Exit Sub
    End If
    Select Case addMode
        Case "overwrite"
            dict.Item(dictKey) = itemValue
        Case "sum"
            dict.Item(dictKey) = dict.Item(dictKey) + CDbl(itemValue)
        Case "count"
            dict.Item(dictKey) = dict.Item(dictKey) + 1
        Case "keep"
            ' the first value stays
        Case Else
            Err.Raise vbObjectError + 518, "ForgeDictAdd", "unknown DICT_ADD mode " & addMode
    End Select
End Sub
