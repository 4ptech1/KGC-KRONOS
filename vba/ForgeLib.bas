' ===== ForgeLib: VCL runtime for Excel 2016 VBA (generated code calls these) =====
' Every function here implements one decision in docs/vcl-semantics.md. Keep them in sync.

Private gForgeLog As Collection
Private gForgeFso As Object

' VCL 0.3 limits: the same numbers as vcl/semantics.py and compiler/vcl_interp.py (a unit test checks)
Private Const FORGE_MAX_NAME As Long = 100
Private Const FORGE_MAX_FILE_PATH As Long = 258
Private Const FORGE_MAX_FOLDER_PATH As Long = 247
Private Const FORGE_NAME_FORBIDDEN As String = "\/:*?""<>|"

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
    ' S17 also for WRITE_ARRAY (D36): text Excel would convert gets the hidden ' prefix, in a copy.
    Dim out As Variant
    Dim r As Long
    Dim c As Long
    out = arr
    For r = LBound(out, 1) To UBound(out, 1)
        For c = LBound(out, 2) To UBound(out, 2)
            If VarType(out(r, c)) = vbString Then
                If ForgeNeedsTextGuard(out(r, c)) Then out(r, c) = "'" & out(r, c)
            End If
        Next c
    Next r
    topLeft.Resize(UBound(out, 1) - LBound(out, 1) + 1, UBound(out, 2) - LBound(out, 2) + 1).Value = out
End Sub

Public Function ForgeReadArray(ByVal rng As Range) As Variant
    ' READ_ARRAY (D36): always a 2-D array from 1, even for one cell, whose .Value is not an array.
    Dim one(1 To 1, 1 To 1) As Variant
    If rng.Cells.CountLarge = 1 Then
        one(1, 1) = rng.Value
        ForgeReadArray = one
    Else
        ForgeReadArray = rng.Value
    End If
End Function

Public Sub ForgeInsertRow(ByVal ws As Worksheet, ByVal rowNo As Long)
    ' INSERT_ROW (D36): a blank row with no fill and no number format, whatever the row above holds.
    ws.Rows(rowNo).Insert
    ws.Rows(rowNo).ClearFormats
End Sub

Public Function ForgeFind(ByVal rng As Range, ByVal x As Variant, ByVal whole As Boolean) As Range
    ' FIND (S22): the first cell, row by row, that is SAME as x (whole) or whose text holds TRIM(x),
    ' case ignored (part); Nothing when none does. It loops itself: Range.Find keeps LookIn, LookAt
    ' and MatchCase from the last search, even one made by hand in Excel.
    Dim cell As Range
    Dim want As String
    want = LCase$(ForgeTrim(x))
    If Not whole And Len(want) = 0 Then Exit Function
    For Each cell In rng.Cells
        If whole Then
            If ForgeSame(cell.Value, x) Then
                Set ForgeFind = cell
                Exit Function
            End If
        ElseIf InStr(1, LCase$(ForgeText(cell.Value)), want, vbBinaryCompare) > 0 Then
            Set ForgeFind = cell
            Exit Function
        End If
    Next cell
End Function

Public Function ForgeFileName(ByVal path As Variant) As String
    ' FILE_NAME (S20): the text after the last backslash; all of it when there is none.
    Dim s As String
    s = ForgeText(path)
    ForgeFileName = Mid$(s, InStrRev(s, "\") + 1)
End Function

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

' ---- S21: DEDUPE keys (VEP 0.4 Phase 2, D38) ----

Public Function ForgeRowKey(ByVal values As Variant) As Variant
    ' The key of a row: each key value as a dictionary key (S10), its length first so the columns
    ' stay apart. Null when a key cell is empty or an error: such a row is never a duplicate.
    Dim i As Long
    Dim k As String
    Dim out As String
    For i = LBound(values) To UBound(values)
        If ForgeIsEmpty(values(i)) Or ForgeIsError(values(i)) Then
            ForgeRowKey = Null
            Exit Function
        End If
        k = ForgeText(values(i))
        out = out & Len(k) & ":" & k
    Next i
    ForgeRowKey = out
End Function

Public Function ForgeSameRowKey(ByVal a As Variant, ByVal b As Variant) As Boolean
    Dim ka As Variant
    Dim kb As Variant
    ka = ForgeRowKey(a)
    kb = ForgeRowKey(b)
    If IsNull(ka) Or IsNull(kb) Then Exit Function
    ForgeSameRowKey = (StrComp(ka, kb, vbTextCompare) = 0)
End Function

Public Sub ForgeDedupe(ByVal rng As Range, ByVal cols As Variant, ByVal hasHeader As Boolean)
    ' DEDUPE: keeps the first row of each key; each later row with that key is deleted as a whole
    ' sheet row, bottom up, so the rows below move up. cols are sheet column numbers, as SORT key= is.
    Dim seen As Object
    Dim doomed() As Boolean
    Dim parts() As Variant
    Dim key As Variant
    Dim first As Long
    Dim r As Long
    Dim i As Long
    Set seen = ForgeNewDict()
    If hasHeader Then first = 2 Else first = 1
    If rng.Rows.Count < first Then Exit Sub
    ReDim doomed(first To rng.Rows.Count)
    ReDim parts(LBound(cols) To UBound(cols))
    For r = first To rng.Rows.Count
        For i = LBound(cols) To UBound(cols)
            parts(i) = rng.Worksheet.Cells(rng.Row + r - 1, cols(i)).Value
        Next i
        key = ForgeRowKey(parts)
        If Not IsNull(key) Then
            If seen.Exists(key) Then doomed(r) = True Else seen.Add key, True
        End If
    Next r
    For r = rng.Rows.Count To first Step -1
        If doomed(r) Then rng.Worksheet.Rows(rng.Row + r - 1).Delete
    Next r
End Sub

' ---- VCL 0.3: S18 SAFE_NAME, S19 PATH_JOIN, FOLDER_EXISTS and the file ops (E11) ----
' One attempt, then a status: the file ops never raise, overwrite, delete or retry.

Private Function ForgeFso() As Object
    ' Late-bound, so no References are needed. Made once per Excel session.
    If gForgeFso Is Nothing Then Set gForgeFso = CreateObject("Scripting.FileSystemObject")
    Set ForgeFso = gForgeFso
End Function

Private Function ForgeNoTrailingSlash(ByVal s As String) As String
    Do While Right$(s, 1) = "\"
        s = Left$(s, Len(s) - 1)
    Loop
    ForgeNoTrailingSlash = s
End Function

Private Function ForgeHasWildcard(ByVal s As String) As Boolean
    ' FSO reads * and ? as patterns, so a path with one is never used.
    ForgeHasWildcard = (InStr(1, s, "*", vbBinaryCompare) > 0 Or InStr(1, s, "?", vbBinaryCompare) > 0)
End Function

Private Function ForgeSamePath(ByVal a As String, ByVal b As String) As Boolean
    ' Paths compare ignoring case, as on Windows.
    ForgeSamePath = (StrComp(LCase$(a), LCase$(b), vbBinaryCompare) = 0)
End Function

Public Function ForgeSafeName(ByVal v As Variant) As String
    Dim s As String
    Dim stem As String
    Dim i As Long
    s = ForgeCleanText(v)
    If Len(s) = 0 Or Len(s) > FORGE_MAX_NAME Then Exit Function
    For i = 1 To Len(s)
        If InStr(1, FORGE_NAME_FORBIDDEN, Mid$(s, i, 1), vbBinaryCompare) > 0 Then Exit Function
    Next i
    ' A trailing dot also covers "." and "..".
    If Right$(s, 1) = "." Then Exit Function
    ' Device names are reserved with any extension and any case: con.txt is CON.
    stem = s
    i = InStr(1, stem, ".", vbBinaryCompare)
    If i > 0 Then stem = Left$(stem, i - 1)
    stem = UCase$(RTrim$(stem))
    If stem = "CON" Or stem = "PRN" Or stem = "AUX" Or stem = "NUL" Then Exit Function
    If Len(stem) = 4 And (Left$(stem, 3) = "COM" Or Left$(stem, 3) = "LPT") Then
        If AscW(Right$(stem, 1)) >= 49 And AscW(Right$(stem, 1)) <= 57 Then Exit Function
    End If
    ForgeSafeName = s
End Function

Public Function ForgePathJoin(ByVal root As Variant, ByVal names As Variant) As String
    ' names is Array(part, ...). An empty or unsafe part gives "", so it never becomes a path.
    Dim s As String
    Dim part As String
    Dim i As Long
    s = ForgeNoTrailingSlash(ForgeText(root))
    If Len(s) = 0 Or UBound(names) < LBound(names) Then Exit Function
    For i = LBound(names) To UBound(names)
        part = ForgeText(names(i))
        If Len(part) = 0 Then Exit Function
        If StrComp(ForgeSafeName(part), part, vbBinaryCompare) <> 0 Then Exit Function
        s = s & "\" & part
    Next i
    ForgePathJoin = s
End Function

Public Function ForgeFolderExists(ByVal path As String) As Boolean
    ' TRUE only for a folder written in full, as in the interpreter's mock. A relative path, or one
    ' Windows would rewrite (a trailing \, a /), depends on the PC, so it gives FALSE.
    On Error GoTo FailedOp
    If Len(path) = 0 Then Exit Function
    If Not ForgeSamePath(ForgeFso().GetAbsolutePathName(path), path) Then Exit Function
    ForgeFolderExists = ForgeFso().FolderExists(path)
    Exit Function
FailedOp:
    ForgeLog "ForgeFolderExists", Err.Number, Err.Description & ": " & path, 0
End Function

Private Function ForgeBelowRoot(ByVal root As String, ByVal path As String) As Boolean
    ' The run-time root check (docs/vep-0.3-file-ops.md, item 3): root is written in full, and
    ' path starts with root and \, as written and as Windows reads it. PATH_JOIN's safe names
    ' keep every generated path below its root; this catches a path built any other way.
    Dim r As String
    r = ForgeNoTrailingSlash(root)
    If Len(r) = 0 Then Exit Function
    If Not ForgeSamePath(ForgeNoTrailingSlash(ForgeFso().GetAbsolutePathName(root)), r) Then Exit Function
    If Not ForgeSamePath(Left$(path, Len(r) + 1), r & "\") Then Exit Function
    ForgeBelowRoot = ForgeSamePath(Left$(ForgeFso().GetAbsolutePathName(path), Len(r) + 1), r & "\")
End Function

Public Function ForgeMakeFolder(ByVal root As String, ByVal path As String) As String
    ' MAKE_FOLDER: path and every missing folder between it and root, never root itself. Every
    ' level is checked before any folder is made, so a blocked path changes nothing.
    Dim fso As Object
    Dim r As String
    Dim parts As Variant
    Dim cur As String
    Dim blocked As Boolean
    Dim i As Long
    On Error GoTo FailedOp
    Set fso = ForgeFso()
    r = ForgeNoTrailingSlash(root)
    If Len(path) = 0 Or Len(r) = 0 Or ForgeHasWildcard(path) Then
        ForgeMakeFolder = "Bad name"
    ElseIf Len(path) > FORGE_MAX_FOLDER_PATH Then
        ForgeMakeFolder = "Too long"
    ElseIf fso.FolderExists(path) Then
        ForgeMakeFolder = "Exists"
    ElseIf fso.FileExists(path) Then
        ForgeMakeFolder = "Blocked by file"
    ElseIf Not ForgeBelowRoot(root, path) Then
        ForgeMakeFolder = "Bad name"
    ElseIf Not fso.FolderExists(r) Then
        ForgeMakeFolder = "No folder"
    Else
        parts = Split(Mid$(path, Len(r) + 2), "\")
        cur = r
        For i = LBound(parts) To UBound(parts)
            cur = cur & "\" & parts(i)
            If fso.FileExists(cur) Then blocked = True
        Next i
        If blocked Then
            ForgeMakeFolder = "Blocked by file"
        Else
            cur = r
            For i = LBound(parts) To UBound(parts)
                cur = cur & "\" & parts(i)
                If Not fso.FolderExists(cur) Then fso.CreateFolder cur
            Next i
            ForgeMakeFolder = "Created"
        End If
    End If
    Exit Function
FailedOp:
    ForgeLog "ForgeMakeFolder", Err.Number, Err.Description & ": " & path, 0
    ForgeMakeFolder = "Failed"
End Function

Public Function ForgeCopyFile(ByVal srcRoot As String, ByVal src As String, ByVal dstRoot As String, ByVal dst As String) As String
    ' COPY_FILE: one copy, never over anything at dst; src always stays, and a short copy is kept.
    Dim fso As Object
    Dim slash As Long
    Dim copying As Boolean
    On Error GoTo FailedOp
    Set fso = ForgeFso()
    slash = InStrRev(dst, "\")
    If Len(src) = 0 Or Len(dst) = 0 Or ForgeHasWildcard(src & dst) Then
        ForgeCopyFile = "Bad name"
    ElseIf Len(src) > FORGE_MAX_FILE_PATH Or Len(dst) > FORGE_MAX_FILE_PATH Then
        ForgeCopyFile = "Too long"
    ElseIf Not ForgeBelowRoot(srcRoot, src) Or Not ForgeBelowRoot(dstRoot, dst) Then
        ForgeCopyFile = "Bad name"
    ElseIf Not fso.FileExists(src) Then
        ForgeCopyFile = "Missing"
    ElseIf slash = 0 Then
        ForgeCopyFile = "No folder"
    ElseIf Not fso.FolderExists(Left$(dst, slash - 1)) Then
        ForgeCopyFile = "No folder"
    ElseIf fso.FileExists(dst) Then
        If fso.GetFile(dst).Size = fso.GetFile(src).Size Then ForgeCopyFile = "Already there" Else ForgeCopyFile = "Conflict"
    ElseIf fso.FolderExists(dst) Then
        ForgeCopyFile = "Conflict"
    Else
        ' overwrite is False: a file that appears at dst meanwhile raises an error, never replaced
        copying = True
        fso.CopyFile src, dst, False
        copying = False
        If fso.GetFile(dst).Size = fso.GetFile(src).Size Then ForgeCopyFile = "Copied" Else ForgeCopyFile = "Check copy"
    End If
    Exit Function
FailedOp:
    ' Error 70 (permission denied) during the copy: another program has src open.
    If copying And Err.Number = 70 Then
        ForgeCopyFile = "Locked"
    Else
        ForgeLog "ForgeCopyFile", Err.Number, Err.Description & ": " & src & " -> " & dst, 0
        ForgeCopyFile = "Failed"
    End If
End Function
