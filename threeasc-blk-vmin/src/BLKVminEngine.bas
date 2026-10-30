Attribute VB_Name = "BLKVminEngine"
Option Explicit
Option Private Module
Private Const READ_BYTES As Long = 262144
Private Const MAX_EXACT As Double = 9007199254740991#
Private Const MAX_LINE As Long = 1048576
Private mTags As Object, mTagHeads As Object, mBlockTag As Object, mSpaces As Object, mDigits As Object
Private mData As Object, mWanted As Object, mDUTs As Object
Private mTno As String, mX As String, mY As String, mPath As String
Private mActive As Boolean, mRecordKey As String, mRecordDUT As String
Private mValues() As Double, mCount As Long, mLine As Long, mRecordLine As Long

Private mSequence() As String, mVoltage() As Double, mConfiguredDUTs As Collection
Private mThreshold As Double, mCarry As String, mBLKError As String

Public Sub BLK_Setup(ByVal cfg As Worksheet)
    ReadConfiguration cfg, mSequence, mVoltage, mConfiguredDUTs, mThreshold
End Sub

Public Sub BLK_BeginFile(ByVal path As String)
    Dim i As Long
    Set mWanted = CreateObject("Scripting.Dictionary")
    For i = LBound(mSequence) To UBound(mSequence): mWanted.Add mSequence(i), True: Next i
    ResetParser path
    mCarry = "": mBLKError = ""
End Sub

Public Sub BLK_Feed(ByVal raw As String, ByVal atEOF As Boolean)
    Dim p As Long, line As String
    If Len(mBLKError) > 0 Then Exit Sub
    On Error GoTo Failed
    mCarry = mCarry & raw
    Do
        p = InStr(mCarry, vbLf)
        If p = 0 Then Exit Do
        mLine = mLine + 1: line = Replace(Left$(mCarry, p - 1), vbCr, "")
        If mActive Or InStr(line, "=") > 0 Then ParseLine line
        mCarry = Mid$(mCarry, p + 1)
    Loop
    If Len(mCarry) > MAX_LINE Then Err.Raise vbObjectError + 310, , "ASC line exceeds 1 MB."
    If atEOF Then
        If Len(mCarry) > 0 Then
            mLine = mLine + 1: line = Replace(mCarry, vbCr, "")
            If mActive Or InStr(line, "=") > 0 Then ParseLine line
        End If
        FinishRecord: mCarry = ""
    End If
    Exit Sub
Failed:
    mBLKError = "FBC Parse Error (" & Err.Number & "): " & Err.Description & "; line=" & mLine
    mCarry = "": mActive = False
End Sub

Public Function BLK_WritePhase(ByVal ws As Worksheet) As String
    Dim duts As Collection, key As Variant, i As Long, block As Long, groupSize As Long, firstRow As Long
    Dim totalRows As Double, header(1 To 1, 1 To 260) As Variant, group As Variant, errorText As String
    On Error GoTo Failed
    ws.UsedRange.UnMerge: ws.UsedRange.Clear
    header(1, 1) = "TNO / Vmin Label": header(1, 2) = "Test Voltage": header(1, 3) = "DUT X": header(1, 4) = "DUT Y"
    For block = 0 To 255: header(1, block + 5) = "BLK " & block: Next block
    ws.Range("A1:IZ1").Value2 = header
    With ws.Range("A1:IZ2")
        .Font.Name = "Calibri": .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 26
    End With
    ws.Columns("A").ColumnWidth = 26: ws.Columns("B").ColumnWidth = 15
    ws.Columns("C:D").ColumnWidth = 10: ws.Columns("E:IZ").ColumnWidth = 11
    ws.Range("A1:IZ1").Interior.Color = RGB(23, 50, 77): ws.Range("A1:IZ1").Font.Color = vbWhite
    If Len(mBLKError) > 0 Then
        ws.Range("A2:D2").Merge: ws.Range("A2").Value2 = mBLKError
        ws.Range("A2:D2").WrapText = True: ws.Rows(2).RowHeight = 60
        BLK_WritePhase = mBLKError: Exit Function
    End If
    Set duts = New Collection
    If mConfiguredDUTs.Count > 0 Then
        For Each key In mConfiguredDUTs: duts.Add CStr(key): Next key
    Else
        Set duts = SortedDUTs(mDUTs)
    End If
    If duts.Count = 0 Then
        ws.Range("A2:D2").Merge: ws.Range("A2").Value2 = "No Data: no FBC BY BLK records for configured TNOs."
        ws.Range("A2:D2").WrapText = True: ws.Rows(2).RowHeight = 60
        BLK_WritePhase = "No selected FBC BY BLK data; Fail Address import preserved.": Exit Function
    End If
    groupSize = UBound(mSequence) + 1: totalRows = CDbl(duts.Count) * groupSize + 1
    If totalRows > ws.Rows.Count Then Err.Raise vbObjectError + 308, , "BLK output exceeds worksheet row limit."
    firstRow = 2
    For Each key In duts
        group = BuildDUTGroup(mData, mSequence, mVoltage, CStr(key), mThreshold)
        WriteGroup ws, firstRow, group: firstRow = firstRow + groupSize
    Next key
    With ws.Range("A1:IZ" & CLng(totalRows))
        .Font.Name = "Calibri": .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 26
    End With
    ws.Columns("A").ColumnWidth = 26: ws.Columns("B").ColumnWidth = 15
    ws.Columns("C:D").ColumnWidth = 10: ws.Columns("E:IZ").ColumnWidth = 11
    ws.Columns("B").NumberFormat = "0.000"
    With ws.Range("A1:IZ1")
        .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: .RowHeight = 32
    End With
    firstRow = 2 + groupSize - 1
    For i = 1 To duts.Count
        ws.Cells(firstRow, 1).Resize(1, 4).Merge
        With ws.Cells(firstRow, 1).Resize(1, 260)
            .Interior.Color = RGB(37, 99, 168): .Font.Color = vbWhite: .Font.Bold = True
        End With
        firstRow = firstRow + groupSize
    Next i
    Exit Function
Failed:
    errorText = "BLK Output Error (" & Err.Number & "): " & Err.Description
    On Error Resume Next
    ws.UsedRange.UnMerge: ws.UsedRange.ClearContents
    ws.Range("A1:IZ1").Value2 = header
    ws.Range("A2:D2").Merge: ws.Range("A2").Value2 = errorText
    BLK_WritePhase = errorText
End Function

Public Sub BLK_Release()
    Set mData = Nothing: Set mWanted = Nothing: Set mDUTs = Nothing
    mCarry = "": mActive = False
End Sub

Private Sub ReadConfiguration(ByVal ws As Worksheet, ByRef tnos() As String, ByRef volts() As Double, ByRef dutList As Collection, ByRef threshold As Double)
    Dim data As Variant, xy As Variant, last As Long, i As Long, count As Long, tno As String, key As String
    Dim seen As Object, duts As Object, value As Variant
    Set seen = CreateObject("Scripting.Dictionary"): Set duts = CreateObject("Scripting.Dictionary")
    Set dutList = New Collection
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last < 2 Then Err.Raise vbObjectError + 301, , "Enter TNO and voltage in Config!A2:B..."
    If last > 901 Then Err.Raise vbObjectError + 319, , "Use Config!A2:B901 for both address and BLK TNOs."
    data = ws.Range("A2:B" & last).Value2
    ReDim tnos(1 To UBound(data, 1)): ReDim volts(1 To UBound(data, 1))
    For i = 1 To UBound(data, 1)
        If IsError(data(i, 1)) Or IsError(data(i, 2)) Then Err.Raise vbObjectError + 302, , "Configuration cell error at row " & i + 1
        If Len(Trim$(CStr(data(i, 1)))) > 0 Then
            tno = CStr(CLng(UnsignedValue(CStr(data(i, 1)), 2147483647#)))
            If CLng(tno) < 100 Or CLng(tno) > 999 Then Err.Raise vbObjectError + 302, , "Configured TNO must be 100-999."
            If seen.Exists(tno) Then Err.Raise vbObjectError + 303, , "Duplicate configured TNO: " & tno
            If Len(Trim$(CStr(data(i, 2)))) = 0 Or Not IsNumeric(data(i, 2)) Or VarType(data(i, 2)) = vbBoolean Then Err.Raise vbObjectError + 304, , "Missing/invalid voltage at Config row " & i + 1
            count = count + 1: tnos(count) = tno: volts(count) = CDbl(data(i, 2)): seen.Add tno, True
        ElseIf Len(Trim$(CStr(data(i, 2)))) > 0 Then
            Err.Raise vbObjectError + 304, , "Voltage without TNO at Config row " & i + 1
        End If
    Next i
    If count = 0 Then Err.Raise vbObjectError + 301, , "No configured TNO."
    ReDim Preserve tnos(1 To count): ReDim Preserve volts(1 To count)
    value = ws.Range("N2").Value2
    If IsError(value) Then Err.Raise vbObjectError + 305, , "Invalid VminFBCThreshold."
    threshold = 0
    If Len(Trim$(CStr(value))) > 0 Then threshold = UnsignedValue(CStr(value), MAX_EXACT)
    last = ws.Cells(ws.Rows.Count, 14).End(xlUp).Row
    i = ws.Cells(ws.Rows.Count, 15).End(xlUp).Row: If i > last Then last = i
    If last >= 6 Then
        xy = ws.Range("N6:O" & last).Value2
        For i = 1 To UBound(xy, 1)
            If IsError(xy(i, 1)) Or IsError(xy(i, 2)) Then Err.Raise vbObjectError + 306, , "Invalid DUT list at row " & i + 5
            If Len(Trim$(CStr(xy(i, 1)))) > 0 Or Len(Trim$(CStr(xy(i, 2)))) > 0 Then
                key = Coordinate(CStr(xy(i, 1))) & "|" & Coordinate(CStr(xy(i, 2)))
                If Not duts.Exists(key) Then dutList.Add key: duts.Add key, True
            End If
        Next i
    End If
End Sub

Private Sub ResetParser(ByVal path As String)
    Set mData = CreateObject("Scripting.Dictionary"): Set mDUTs = CreateObject("Scripting.Dictionary")
    mTno = "": mX = "": mY = "": mPath = path: mLine = 0
    mActive = False: mCount = 0
    If mTags Is Nothing Then
        Set mTags = CreateObject("VBScript.RegExp"): mTags.Global = True: mTags.IgnoreCase = True
        mTags.Pattern = "\b(Tno|X|Y)\s*=\s*([+-]?[0-9]+)(?=\s|[,;]|$)"
        Set mTagHeads = CreateObject("VBScript.RegExp"): mTagHeads.Global = True: mTagHeads.IgnoreCase = True
        mTagHeads.Pattern = "\b(Tno|X|Y)\s*="
        Set mBlockTag = CreateObject("VBScript.RegExp"): mBlockTag.IgnoreCase = True
        mBlockTag.Pattern = "\bFBC\s+BY\s+BLK\s*=\s*"
        Set mSpaces = CreateObject("VBScript.RegExp"): mSpaces.Global = True: mSpaces.Pattern = "[\s,;]+"
        Set mDigits = CreateObject("VBScript.RegExp"): mDigits.Pattern = "^[0-9]+$"
    End If
End Sub

Private Sub ParseLine(ByVal line As String)
    Dim blocks As Object, tags As Object, tag As Object, before As String, payload As String
    Dim name As String, value As String, text As String
    Set blocks = mBlockTag.Execute(line)
    before = line
    If blocks.Count > 0 Then before = Left$(line, blocks(0).FirstIndex)
    Set tags = mTags.Execute(before)
    If mTagHeads.Execute(before).Count <> tags.Count Then Err.Raise vbObjectError + 318, , "Malformed TNO/X/Y header; refusing stale context."
    If blocks.Count > 0 Or tags.Count > 0 Then
        FinishRecord
    ElseIf mActive Then
        text = Trim$(line)
        If Len(text) = 0 Then Exit Sub
        If mCount = 256 And InStr(1, "0123456789,+-", Left$(text, 1), vbBinaryCompare) = 0 Then
            FinishRecord
        Else
            AddValues line: Exit Sub
        End If
    End If
    For Each tag In tags
        name = UCase$(tag.SubMatches(0)): value = CStr(tag.SubMatches(1))
        Select Case name
            Case "TNO": mTno = CStr(CLng(UnsignedValue(value, 2147483647#)))
                If mTno = "0" Then Err.Raise vbObjectError + 311, , "TNO must be positive."
            Case "X": mX = Coordinate(value)
            Case "Y": mY = Coordinate(value)
        End Select
    Next tag
    If blocks.Count > 0 Then
        If Len(mTno) = 0 Or Len(mX) = 0 Or Len(mY) = 0 Then Err.Raise vbObjectError + 312, , "FBC record has no complete TNO/X/Y context."
        mRecordDUT = mX & "|" & mY: mRecordKey = mTno & "|" & mRecordDUT
        mRecordLine = mLine: mActive = True: mCount = 0: ReDim mValues(0 To 255)
        ' FirstIndex is zero based; Length includes '=', so +1 preserves the first value.
        payload = Mid$(line, blocks(0).FirstIndex + blocks(0).Length + 1)
        AddValues payload
    End If
End Sub

Private Sub AddValues(ByVal payload As String)
    Dim tokens As Variant, token As Variant
    payload = Trim$(mSpaces.Replace(payload, " "))
    If Len(payload) = 0 Then Exit Sub
    tokens = Split(payload, " ")
    For Each token In tokens
        If mCount >= 256 Then Err.Raise vbObjectError + 313, , "FBC record exceeds 256 values; record line=" & mRecordLine
        mValues(mCount) = UnsignedValue(CStr(token), MAX_EXACT)
        mCount = mCount + 1
    Next token
End Sub

Private Sub FinishRecord()
    Dim tno As String
    If Not mActive Then Exit Sub
    If mCount <> 256 Then Err.Raise vbObjectError + 314, , "FBC record requires exactly 256 values; found " & mCount & "; record line=" & mRecordLine
    tno = Split(mRecordKey, "|")(0)
    If mWanted.Exists(tno) Then
        If Not mData.Exists(mRecordKey) Then mData.Add mRecordKey, mValues
        If Not mDUTs.Exists(mRecordDUT) Then mDUTs.Add mRecordDUT, True
    End If
    mActive = False
End Sub

Private Function BuildDUTGroup(ByVal data As Object, ByRef tnos() As String, ByRef volts() As Double, ByVal dut As String, ByVal threshold As Double) As Variant
    Dim output() As Variant, cache() As Variant, present() As Boolean, xy As Variant
    Dim n As Long, i As Long, block As Long, anyData As Boolean, key As String
    n = UBound(tnos): ReDim output(1 To n + 1, 1 To 260)
    ReDim cache(1 To n): ReDim present(1 To n): xy = Split(dut, "|")
    For i = 1 To n
        output(i, 1) = CLng(tnos(i)): output(i, 2) = volts(i)
        output(i, 3) = CLng(xy(0)): output(i, 4) = CLng(xy(1))
        key = tnos(i) & "|" & dut: present(i) = data.Exists(key)
        If present(i) Then cache(i) = data(key): anyData = True
        For block = 0 To 255
            If present(i) Then output(i, block + 5) = cache(i)(block) Else output(i, block + 5) = "Not Found"
        Next block
    Next i
    output(n + 1, 1) = "X" & xy(0) & "Y" & xy(1) & " Vmin"
    For block = 0 To 255
        If anyData Then output(n + 1, block + 5) = JudgeBlock(cache, present, volts, block, threshold) Else output(n + 1, block + 5) = "No Data"
    Next block
    BuildDUTGroup = output
End Function

Private Function JudgeBlock(ByRef cache() As Variant, ByRef present() As Boolean, ByRef volts() As Double, ByVal block As Long, ByVal threshold As Double) As Variant
    Dim i As Long, cumulative As Double, value As Double, trigger As Boolean
    For i = LBound(present) To UBound(present)
        If Not present(i) Then JudgeBlock = "Missing Data": Exit Function
        value = CDbl(cache(i)(block))
        If threshold = 0 Then
            trigger = (value <> 0)
        Else
            ' Compare before adding to avoid exceeding exact integer precision.
            trigger = (value >= threshold - cumulative)
            If Not trigger Then cumulative = cumulative + value
        End If
        If trigger Then
            If i = LBound(present) Then JudgeBlock = "Fail" Else JudgeBlock = volts(i - 1)
            Exit Function
        End If
    Next i
    JudgeBlock = "Pass"
End Function

Private Function SortedDUTs(ByVal duts As Object) As Collection
    Dim result As New Collection, keys As Variant, key As Variant
    If duts.Count > 0 Then
        keys = duts.Keys: SortXY keys, LBound(keys), UBound(keys)
        For Each key In keys: result.Add CStr(key): Next key
    End If
    Set SortedDUTs = result
End Function

Private Sub SortXY(ByRef keys As Variant, ByVal lo As Long, ByVal hi As Long)
    Dim left As Long, right As Long, pivot As String, swap As Variant
    left = lo: right = hi: pivot = CStr(keys((lo + hi) \ 2))
    Do While left <= right
        Do While CompareXY(CStr(keys(left)), pivot) < 0: left = left + 1: Loop
        Do While CompareXY(CStr(keys(right)), pivot) > 0: right = right - 1: Loop
        If left <= right Then
            swap = keys(left): keys(left) = keys(right): keys(right) = swap
            left = left + 1: right = right - 1
        End If
    Loop
    If lo < right Then SortXY keys, lo, right
    If left < hi Then SortXY keys, left, hi
End Sub

Private Function CompareXY(ByVal a As String, ByVal b As String) As Long
    Dim aa As Variant, bb As Variant
    aa = Split(a, "|"): bb = Split(b, "|")
    If CDbl(aa(0)) < CDbl(bb(0)) Then CompareXY = -1: Exit Function
    If CDbl(aa(0)) > CDbl(bb(0)) Then CompareXY = 1: Exit Function
    If CDbl(aa(1)) < CDbl(bb(1)) Then CompareXY = -1: Exit Function
    If CDbl(aa(1)) > CDbl(bb(1)) Then CompareXY = 1
End Function

Private Function UnsignedValue(ByVal text As String, ByVal maximum As Double) As Double
    Dim re As Object
    text = Trim$(text)
    If mDigits Is Nothing Then Set mDigits = CreateObject("VBScript.RegExp"): mDigits.Pattern = "^[0-9]+$"
    If Not mDigits.Test(text) Then Err.Raise vbObjectError + 315, , "Expected non-negative integer, found: " & Left$(text, 80)
    If Len(text) > 1024 Then Err.Raise vbObjectError + 315, , "Integer token too long."
    If CDbl(text) > maximum Then Err.Raise vbObjectError + 315, , "Integer exceeds supported exact range."
    UnsignedValue = CDbl(text)
End Function

Private Function Coordinate(ByVal text As String) As String
    Dim re As Object, value As Double
    text = Trim$(text): Set re = CreateObject("VBScript.RegExp"): re.Pattern = "^[+-]?[0-9]+$"
    If Not re.Test(text) Then Err.Raise vbObjectError + 316, , "Invalid DUT coordinate: " & text
    value = CDbl(text)
    If value < -2147483648# Or value > 2147483647# Then Err.Raise vbObjectError + 316, , "DUT coordinate out of range."
    Coordinate = CStr(CLng(value))
End Function

Private Sub WriteGroup(ByVal ws As Worksheet, ByVal firstRow As Long, ByRef group As Variant)
    Dim offset As Long, take As Long, r As Long, c As Long, packet() As Variant
    offset = 1
    Do While offset <= UBound(group, 1)
        take = UBound(group, 1) - offset + 1: If take > 504 Then take = 504
        ReDim packet(1 To take, 1 To 260)
        For r = 1 To take
            For c = 1 To 260: packet(r, c) = group(offset + r - 1, c): Next c
        Next r
        ws.Cells(firstRow + offset - 1, 1).Resize(take, 260).Value2 = packet
        offset = offset + take
    Loop
End Sub

