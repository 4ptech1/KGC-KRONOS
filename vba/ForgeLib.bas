' ===== ForgeLib: VCL runtime for Excel 2016 VBA (generated code calls these) =====
' Every function here implements one decision in docs/vcl-semantics.md. Keep them in sync.

Private gForgeLog As Collection

Public Sub ForgeLog(ByVal procName As String, ByVal errNumber As Long, ByVal message As String, ByVal lineNo As Long)
    If gForgeLog Is Nothing Then Set gForgeLog = New Collection
    gForgeLog.Add procName & " | error " & errNumber & " | spec line " & (lineNo \ 10) & " | " & message
    Debug.Print "ForgeLog: " & procName & " | " & errNumber & " | " & message
End Sub

Public Sub ForgeInfo(ByVal procName As String, ByVal message As String)
    Debug.Print "ForgeInfo: " & procName & " | " & message
End Sub

Public Function ForgeErrorCount() As Long
    If gForgeLog Is Nothing Then ForgeErrorCount = 0 Else ForgeErrorCount = gForgeLog.Count
End Function

Public Sub ForgeClearLog()
    Set gForgeLog = New Collection
End Sub

Public Function ForgeLogText() As String
    Dim item As Variant
    Dim s As String
    If gForgeLog Is Nothing Then Exit Function
    For Each item In gForgeLog
        s = s & item & vbCrLf
    Next item
    ForgeLogText = s
End Function

Public Sub ForgeRestore(ByVal screenOn As Boolean, ByVal calcMode As Long, ByVal eventsOn As Boolean)
    On Error Resume Next
    Application.ScreenUpdating = screenOn
    Application.Calculation = calcMode
    Application.EnableEvents = eventsOn
    On Error GoTo 0
End Sub

' ---- S1 to S4: emptiness, numbers, dates ----

Private Function ForgeIsSpaceChar(ByVal code As Long) As Boolean
    ForgeIsSpaceChar = (code = 32 Or code = 160 Or code = 9 Or code = 10 Or code = 13)
End Function

Public Function ForgeIsEmpty(ByVal v As Variant) As Boolean
    Dim i As Long
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then
        ForgeIsEmpty = True
        Exit Function
    End If
    If VarType(v) <> vbString Then Exit Function
    For i = 1 To Len(v)
        If Not ForgeIsSpaceChar(AscW(Mid$(v, i, 1))) Then Exit Function
    Next i
    ForgeIsEmpty = True
End Function

Private Function ForgeIsPlainNumberText(ByVal s As String) As Boolean
    Dim i As Long
    Dim ch As String
    Dim digits As Long
    Dim dots As Long
    Dim start As Long
    If Len(s) = 0 Then Exit Function
    start = 1
    If Left$(s, 1) = "+" Or Left$(s, 1) = "-" Then start = 2
    If start > Len(s) Then Exit Function
    For i = start To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits + 1
        ElseIf ch = "." Then
            dots = dots + 1
            If dots > 1 Then Exit Function
        Else
            Exit Function
        End If
    Next i
    If digits = 0 Then Exit Function
    If Right$(s, 1) = "." Then Exit Function
    ForgeIsPlainNumberText = True
End Function

Public Function ForgeIsNumber(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal, vbByte
            ForgeIsNumber = True
        Case vbString
            ForgeIsNumber = ForgeIsPlainNumberText(ForgeTrim(v))
        Case Else
            ForgeIsNumber = False
    End Select
End Function

Public Function ForgeNum(ByVal v As Variant) As Double
    ' Only call after ForgeIsNumber is true. Val always uses "." as the decimal point.
    If VarType(v) = vbString Then
        ForgeNum = Val(ForgeTrim(v))
    Else
        ForgeNum = CDbl(v)
    End If
End Function

Public Function ForgeIsDate(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    ForgeIsDate = (VarType(v) = vbDate)
End Function

' ---- S6: canonical text ----

Private Function ForgeErrorText(ByVal v As Variant) As String
    If v = CVErr(xlErrNA) Then
        ForgeErrorText = "#N/A"
    ElseIf v = CVErr(xlErrDiv0) Then
        ForgeErrorText = "#DIV/0!"
    ElseIf v = CVErr(xlErrValue) Then
        ForgeErrorText = "#VALUE!"
    ElseIf v = CVErr(xlErrRef) Then
        ForgeErrorText = "#REF!"
    ElseIf v = CVErr(xlErrName) Then
        ForgeErrorText = "#NAME?"
    ElseIf v = CVErr(xlErrNum) Then
        ForgeErrorText = "#NUM!"
    ElseIf v = CVErr(xlErrNull) Then
        ForgeErrorText = "#NULL!"
    Else
        ForgeErrorText = "#ERROR"
    End If
End Function

Private Function ForgeTwo(ByVal n As Long) As String
    ForgeTwo = Right$("0" & CStr(n), 2)
End Function

Private Function ForgeNumberText(ByVal d As Double) As String
    Dim s As String
    If d = Int(d) And Abs(d) < 1E+15 Then
        ForgeNumberText = Format$(d, "0")
        Exit Function
    End If
    s = Trim$(Str$(d))
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    ForgeNumberText = s
End Function

Public Function ForgeText(ByVal v As Variant) As String
    Dim d As Date
    If IsError(v) Then
        ForgeText = ForgeErrorText(v)
        Exit Function
    End If
    Select Case VarType(v)
        Case vbEmpty, vbNull
            ForgeText = ""
        Case vbString
            ForgeText = v
        Case vbBoolean
            If v Then ForgeText = "TRUE" Else ForgeText = "FALSE"
        Case vbDate
            d = v
            ForgeText = CStr(Year(d)) & "-" & ForgeTwo(Month(d)) & "-" & ForgeTwo(Day(d))
            If d <> Int(d) Then
                ForgeText = ForgeText & " " & ForgeTwo(Hour(d)) & ":" & ForgeTwo(Minute(d)) & ":" & ForgeTwo(Second(d))
            End If
        Case vbInteger, vbLong, vbByte
            ForgeText = CStr(v)
        Case vbSingle, vbDouble, vbCurrency, vbDecimal
            ForgeText = ForgeNumberText(CDbl(v))
        Case Else
            ForgeText = CStr(v)
    End Select
End Function

' ---- S2: TRIM (also removes non-breaking spaces, tabs and line breaks) ----

Public Function ForgeTrim(ByVal v As Variant) As String
    Dim s As String
    Dim a As Long
    Dim b As Long
    s = ForgeText(v)
    a = 1
    b = Len(s)
    Do While a <= b
        If Not ForgeIsSpaceChar(AscW(Mid$(s, a, 1))) Then Exit Do
        a = a + 1
    Loop
    Do While b >= a
        If Not ForgeIsSpaceChar(AscW(Mid$(s, b, 1))) Then Exit Do
        b = b - 1
    Loop
    If b >= a Then ForgeTrim = Mid$(s, a, b - a + 1) Else ForgeTrim = ""
End Function

Public Function ForgeLeft(ByVal v As Variant, ByVal n As Double) As String
    If n < 0 Then n = 0
    ForgeLeft = Left$(ForgeText(v), CLng(Int(n)))
End Function

Public Function ForgeRight(ByVal v As Variant, ByVal n As Double) As String
    If n < 0 Then n = 0
    ForgeRight = Right$(ForgeText(v), CLng(Int(n)))
End Function

' ---- S7: = between a value and text compares canonical text ----

Public Function ForgeEq(ByVal a As Variant, ByVal b As Variant) As Boolean
    If VarType(a) = vbString Or VarType(b) = vbString Or IsEmpty(a) Or IsEmpty(b) Then
        ForgeEq = (StrComp(ForgeText(a), ForgeText(b), vbBinaryCompare) = 0)
    ElseIf IsError(a) Or IsError(b) Then
        ForgeEq = (ForgeText(a) = ForgeText(b))
    Else
        ForgeEq = (a = b)
    End If
End Function

' ---- VCL 0.2 value functions: S5 and S12 to S16 ----

Public Function ForgeIsError(ByVal v As Variant) As Boolean
    ForgeIsError = IsError(v)
End Function

Public Function ForgeSame(ByVal a As Variant, ByVal b As Variant) As Boolean
    ' Binary compare of the lower-case texts: vbTextCompare would follow the PC's locale.
    ForgeSame = (StrComp(LCase$(ForgeTrim(a)), LCase$(ForgeTrim(b)), vbBinaryCompare) = 0)
End Function

Public Function ForgeCleanText(ByVal v As Variant) As String
    Dim s As String
    Dim ch As String
    Dim code As Long
    Dim i As Long
    Dim pending As Boolean
    Dim out As String
    s = ForgeText(v)
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If ForgeIsSpaceChar(code) Then
            pending = True
        ElseIf code >= 32 Or code < 0 Then
            ' AscW is below 0 for characters above &H7FFF, which are kept
            If pending And Len(out) > 0 Then out = out & " "
            pending = False
            out = out & ch
        End If
    Next i
    ForgeCleanText = out
End Function

Private Function ForgeAllDigits(ByVal s As String, ByVal fewest As Long, ByVal most As Long) As Boolean
    Dim i As Long
    Dim ch As String
    If Len(s) < fewest Or Len(s) > most Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    ForgeAllDigits = True
End Function

Public Function ForgeParseDate(ByVal v As Variant, ByVal fmt As String) As Variant
    ' Never CDate or DateValue: those read the text by the PC's regional settings.
    Dim sep As String
    Dim parts As Variant
    Dim names As Variant
    Dim i As Long
    Dim dd As String
    Dim mm As String
    Dim yyyy As String
    Dim result As Date
    ForgeParseDate = Empty
    If IsError(v) Then Exit Function
    If VarType(v) = vbDate Then
        ForgeParseDate = DateSerial(Year(v), Month(v), Day(v))
        Exit Function
    End If
    If VarType(v) <> vbString Then Exit Function
    If Left$(fmt, 4) = "yyyy" Then sep = Mid$(fmt, 5, 1) Else sep = Mid$(fmt, 3, 1)
    parts = Split(ForgeTrim(v), sep)
    names = Split(fmt, sep)
    If UBound(parts) <> 2 Or UBound(names) <> 2 Then Exit Function
    For i = 0 To 2
        If names(i) = "dd" Then
            dd = parts(i)
        ElseIf names(i) = "mm" Then
            mm = parts(i)
        ElseIf names(i) = "yyyy" Then
            yyyy = parts(i)
        End If
    Next i
    If Not ForgeAllDigits(dd, 1, 2) Then Exit Function
    If Not ForgeAllDigits(mm, 1, 2) Then Exit Function
    If Not ForgeAllDigits(yyyy, 4, 4) Then Exit Function
    If CLng(yyyy) < 1900 Then Exit Function
    If CLng(mm) < 1 Or CLng(mm) > 12 Or CLng(dd) < 1 Then Exit Function
    ' DateSerial rolls 31/02 over into March, so the day must come back unchanged
    result = DateSerial(CLng(yyyy), CLng(mm), CLng(dd))
    If Day(result) <> CLng(dd) Then Exit Function
    ForgeParseDate = result
End Function

Public Function ForgeIif(ByVal cond As Boolean, ByVal a As Variant, ByVal b As Variant) As Variant
    ' Both values were worked out by the caller already, as with VBA's IIf.
    If cond Then ForgeIif = a Else ForgeIif = b
End Function

Public Function ForgeMod(ByVal a As Double, ByVal b As Double) As Double
    ' Excel's MOD: the sign of b, decimals kept. VBA's Mod rounds first. b = 0 raises error 11.
    ForgeMod = a - b * Int(a / b)
End Function

' ---- S9: COL matches headers ignoring case, extra spaces and line breaks ----

Public Function ForgeHeaderKey(ByVal v As Variant) As String
    Dim s As String
    s = ForgeText(v)
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    ForgeHeaderKey = LCase$(Trim$(s))
End Function

Public Function ForgeCol(ByVal ws As Worksheet, ByVal header As Variant) As Long
    Dim c As Long
    Dim lastCol As Long
    Dim want As String
    If ws Is Nothing Then Exit Function
    want = ForgeHeaderKey(header)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If ForgeHeaderKey(ws.Cells(1, c).Value) = want Then
            ForgeCol = c
            Exit Function
        End If
    Next c
End Function

' ---- S11: last row / last column like End(xlUp) ----

Public Function ForgeLastRow(ByVal ws As Worksheet, ByVal col As Long) As Long
    ForgeLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

Public Function ForgeLastCol(ByVal ws As Worksheet, ByVal rowNo As Long) As Long
    ForgeLastCol = ws.Cells(rowNo, ws.Columns.Count).End(xlToLeft).Column
End Function

' ---- ranges ----

Public Function ForgeUsedRange(ByVal ws As Worksheet) As Range
    Set ForgeUsedRange = ws.Range(ws.Cells(1, 1), ws.Cells(ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1, _
        ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1))
End Function

Public Function ForgeRowRange(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal c1 As Long, ByVal c2 As Long) As Range
    Set ForgeRowRange = ws.Range(ws.Cells(rowNo, c1), ws.Cells(rowNo, c2))
End Function

' ---- S17: SET_VALUE, WRITE_ROW and COPY_ROW store text as exactly that text ----

Public Function ForgeNeedsTextGuard(ByVal s As String) As Boolean
    ' True when Excel could turn the text into a number, date, TRUE/FALSE, error or formula.
    Dim i As Long
    Dim code As Long
    Dim t As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1))
        If (code >= 48 And code <= 57) Or code > 127 Or code < 0 Then
            ForgeNeedsTextGuard = True
            Exit Function
        End If
    Next i
    If InStr(1, "=+-@'#", Left$(s, 1), vbBinaryCompare) > 0 Then
        ForgeNeedsTextGuard = True
        Exit Function
    End If
    t = UCase$(ForgeTrim(s))
    ForgeNeedsTextGuard = (t = "TRUE" Or t = "FALSE")
End Function

Public Sub ForgeSetValue(ByVal target As Range, ByVal v As Variant)
    ' The ' prefix is hidden by Excel: the cell holds the text unchanged.
    If VarType(v) = vbString Then
        If ForgeNeedsTextGuard(v) Then
            target.Value = "'" & v
            Exit Sub
        End If
    End If
    target.Value = v
End Sub

Public Sub ForgeWriteRow(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal firstCol As Long, ByVal values As Variant)
    Dim i As Long
    For i = LBound(values) To UBound(values)
        ForgeSetValue ws.Cells(rowNo, firstCol + i - LBound(values)), values(i)
    Next i
End Sub

Public Sub ForgeCopyRow(ByVal src As Worksheet, ByVal srcRow As Long, ByVal dst As Worksheet, ByVal dstRow As Long, ByVal nCols As Long)
    Dim c As Long
    For c = 1 To nCols
        ForgeSetValue dst.Cells(dstRow, c), src.Cells(srcRow, c).Value
    Next c
End Sub

Public Sub ForgeNoFill(ByVal rng As Range)
    rng.Interior.Pattern = xlNone
End Sub

Public Sub ForgeSort(ByVal rng As Range, ByVal keyCol As Long, ByVal ascending As Boolean, ByVal hasHeader As Boolean)
    Dim minRows As Long
    If hasHeader Then minRows = 3 Else minRows = 2
    If rng.Rows.Count < minRows Then Exit Sub
    If ascending Then
        rng.Sort Key1:=rng.Worksheet.Cells(rng.Row, keyCol), Order1:=xlAscending, _
            Header:=IIf(hasHeader, xlYes, xlNo)
    Else
        rng.Sort Key1:=rng.Worksheet.Cells(rng.Row, keyCol), Order1:=xlDescending, _
            Header:=IIf(hasHeader, xlYes, xlNo)
    End If
End Sub

Public Sub ForgeWriteArray(ByVal arr As Variant, ByVal topLeft As Range)
    topLeft.Resize(UBound(arr, 1) - LBound(arr, 1) + 1, UBound(arr, 2) - LBound(arr, 2) + 1).Value = arr
End Sub

' ---- S10: dictionaries with canonical-text keys, case ignored ----

Public Function ForgeNewDict() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    Set ForgeNewDict = d
End Function

Public Sub ForgeDictAdd(ByVal d As Object, ByVal key As Variant, ByVal value As Variant, ByVal mode As String)
    Dim k As String
    k = ForgeText(key)
    Select Case mode
        Case "overwrite"
            d(k) = value
        Case "keep"
            If Not d.Exists(k) Then d.Add k, value
        Case "sum"
            If d.Exists(k) Then
                d(k) = d(k) + ForgeNum(value)
            Else
                d.Add k, ForgeNum(value)
            End If
        Case "count"
            If d.Exists(k) Then
                d(k) = d(k) + 1
            Else
                d.Add k, 1
            End If
    End Select
End Sub

Public Function ForgeDictGet(ByVal d As Object, ByVal key As Variant) As Variant
    Dim k As String
    k = ForgeText(key)
    If d.Exists(k) Then ForgeDictGet = d(k) Else ForgeDictGet = Empty
End Function

Public Function ForgeDictExists(ByVal d As Object, ByVal key As Variant) As Boolean
    ForgeDictExists = d.Exists(ForgeText(key))
End Function
