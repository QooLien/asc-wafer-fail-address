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
Private mTestNames() As String, mGroups As Object, mSheetNames As Object, mNextRows As Object
Private mMonitorTNOs As Object
Private mThreshold As Double, mCarry As String, mBLKError As String

Public Sub BLK_Setup(ByVal cfg As Worksheet)
    ReadConfiguration cfg, mSequence, mTestNames, mVoltage, mConfiguredDUTs, mThreshold
    Set mGroups = BuildNameGroups(mTestNames)
    SetupNameSheets cfg.Parent
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

Public Function BLK_WritePhase(ByVal wb As Workbook, ByVal phase As Long) As String
    Dim name As Variant, warning As String, text As String, warningCount As Long
    For Each name In mGroups
        warning = WriteNamedPhase(wb, CStr(name), phase)
        If Len(warning) > 0 Then
            warningCount = warningCount + 1
            If warningCount <= 5 Then text = text & IIf(Len(text) > 0, "; ", "") & warning
        End If
    Next name
    If warningCount > 5 Then text = text & "; additional Test Names with messages: " & warningCount - 5
    BLK_WritePhase = text
End Function

Private Function WriteNamedPhase(ByVal wb As Workbook, ByVal name As String, ByVal phase As Long) As String
    Dim ws As Worksheet, tnos() As String, volts() As Double, duts As Collection, dut As Variant
    Dim header(1 To 1, 1 To 260) As Variant, group As Variant, firstRow As Long, dataRow As Long
    Dim groupSize As Long, endRow As Long, i As Long, block As Long, color As Long, text As String
    On Error GoTo Failed
    GroupSequence mSequence, mVoltage, mTestNames, name, tnos, volts
    On Error Resume Next
    Set ws = wb.Worksheets(CStr(mSheetNames(name)))
    On Error GoTo Failed
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = CStr(mSheetNames(name))
    End If
    header(1, 1) = "TNO / Vmin Label": header(1, 2) = "Test Voltage": header(1, 3) = "DUT X": header(1, 4) = "DUT Y"
    For block = 0 To 255: header(1, block + 5) = "BLK " & block: Next block
    If phase = 1 Then
        ws.UsedRange.UnMerge: ws.UsedRange.Clear
        ws.Range("A1:IZ1").Value2 = header
        With ws.Range("A1:IZ1")
            .Font.Name = "Calibri": .Font.Size = 11
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: .RowHeight = 32
        End With
        ws.Columns("A").ColumnWidth = 26: ws.Columns("B").ColumnWidth = 15
        ws.Columns("C:D").ColumnWidth = 10: ws.Columns("E:IZ").ColumnWidth = 11
        ws.Columns("B").NumberFormat = "0.000": mNextRows(name) = 2
    End If
    firstRow = CLng(mNextRows(name)): dataRow = firstRow + 1
    Select Case phase
        Case 1: color = RGB(219, 234, 254)
        Case 2: color = RGB(255, 237, 213)
        Case 3: color = RGB(220, 252, 231)
    End Select
    ws.Cells(firstRow, 1).Resize(1, 4).Merge
    ws.Cells(firstRow, 1).Value2 = "WT" & phase & " | " & name
    With ws.Cells(firstRow, 1).Resize(1, 260)
        .Font.Name = "Calibri": .Font.Size = 11: .Font.Bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .Interior.Color = color: .RowHeight = 32
    End With
    ws.Cells(firstRow, 1).WrapText = True
    If Len(mBLKError) > 0 Then text = mBLKError: GoTo NoteOnly
    Set duts = GroupDUTs(tnos)
    If duts.Count = 0 Then text = "No Data: no FBC records for this Test Name.": GoTo NoteOnly
    groupSize = UBound(tnos) + IIf(HasMonitor(tnos, mMonitorTNOs), 1, 0)
    If CDbl(dataRow) + CDbl(duts.Count) * groupSize - 1 > ws.Rows.Count Then Err.Raise vbObjectError + 404, , "Test Name sheet exceeds row limit."
    endRow = dataRow
    For Each dut In duts
        group = BuildItemDUTGroup(mData, tnos, volts, CStr(dut), mThreshold, mMonitorTNOs)
        WriteGroup ws, endRow, group: endRow = endRow + groupSize
    Next dut
    With ws.Cells(dataRow, 1).Resize(endRow - dataRow, 260)
        .Font.Name = "Calibri": .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 26
    End With
    If HasMonitor(tnos, mMonitorTNOs) Then
    For i = 1 To duts.Count
        block = dataRow + i * groupSize - 1
        ws.Cells(block, 1).Resize(1, 4).Merge
        With ws.Cells(block, 1).Resize(1, 260)
            .Interior.Color = RGB(37, 99, 168): .Font.Color = vbWhite: .Font.Bold = True
        End With
    Next i
    End If
    mNextRows(name) = endRow + 1
    Exit Function
NoteOnly:
    ws.Cells(dataRow, 1).Resize(1, 4).Merge: ws.Cells(dataRow, 1).Value2 = text
    With ws.Cells(dataRow, 1).Resize(1, 4)
        .Font.Name = "Calibri": .Font.Size = 11: .WrapText = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 60
    End With
    mNextRows(name) = dataRow + 2
    WriteNamedPhase = name & ": " & text
    Exit Function
Failed:
    text = "BLK Output Error (" & Err.Number & "): " & Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then
        If dataRow > 0 Then
            endRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
            If endRow < dataRow Then endRow = dataRow
            ws.Cells(dataRow, 1).Resize(endRow - dataRow + 1, 260).UnMerge
            ws.Cells(dataRow, 1).Resize(endRow - dataRow + 1, 260).ClearContents
            ws.Cells(dataRow, 1).Resize(1, 4).Merge: ws.Cells(dataRow, 1).Value2 = text
            mNextRows(name) = dataRow + 2
        End If
    End If
    WriteNamedPhase = name & ": " & text
End Function

Public Sub BLK_Release()
    Set mData = Nothing: Set mWanted = Nothing: Set mDUTs = Nothing
    mCarry = "": mActive = False
    Set mGroups = Nothing: Set mSheetNames = Nothing: Set mNextRows = Nothing
End Sub

Private Sub ReadConfiguration(ByVal ws As Worksheet, ByRef tnos() As String, ByRef names() As String, ByRef volts() As Double, ByRef dutList As Collection, ByRef threshold As Double)
    Dim data As Variant, xy As Variant, last As Long, i As Long, count As Long, tno As String, key As String
    Dim seen As Object, duts As Object, value As Variant
    Set seen = CreateObject("Scripting.Dictionary"): Set duts = CreateObject("Scripting.Dictionary")
    Set dutList = New Collection
    Set mMonitorTNOs = CreateObject("Scripting.Dictionary")
    ReDim tnos(1 To 900): ReDim names(1 To 900): ReDim volts(1 To 900)
    ReadItemColumns ws.Range("A2:C901").Value2, True, seen, tnos, names, volts, count
    ReadItemColumns ws.Range("D2:F901").Value2, False, seen, tnos, names, volts, count
    If count = 0 Then Err.Raise vbObjectError + 301, , "Enter Monitor or Function TNOs in Config!A:F."
    ReDim Preserve tnos(1 To count): ReDim Preserve names(1 To count): ReDim Preserve volts(1 To count)
    value = ws.Range("Q2").Value2
    If IsError(value) Then Err.Raise vbObjectError + 305, , "Invalid VminFBCThreshold."
    threshold = 0
    If Len(Trim$(CStr(value))) > 0 Then threshold = UnsignedValue(CStr(value), MAX_EXACT)
    last = ws.Cells(ws.Rows.Count, 17).End(xlUp).Row
    i = ws.Cells(ws.Rows.Count, 18).End(xlUp).Row: If i > last Then last = i
    If last >= 6 Then
        xy = ws.Range("Q6:R" & last).Value2
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


Private Function BuildNameGroups(ByRef names() As String) As Object
    Dim result As Object, i As Long, name As String
    Set result = CreateObject("Scripting.Dictionary"): result.CompareMode = vbTextCompare
    For i = LBound(names) To UBound(names)
        name = Trim$(names(i))
        If Len(name) = 0 Then Err.Raise vbObjectError + 401, , "Test Name cannot be blank."
        If Not result.Exists(name) Then result.Add name, True
    Next i
    Set BuildNameGroups = result
End Function

Private Sub GroupSequence(ByRef allTNOs() As String, ByRef allVolts() As Double, ByRef names() As String, ByVal name As String, ByRef tnos() As String, ByRef volts() As Double)
    Dim i As Long, count As Long
    ReDim tnos(1 To UBound(allTNOs)): ReDim volts(1 To UBound(allTNOs))
    For i = LBound(allTNOs) To UBound(allTNOs)
        If StrComp(Trim$(names(i)), Trim$(name), vbTextCompare) = 0 Then
            count = count + 1: tnos(count) = allTNOs(i): volts(count) = allVolts(i)
        End If
    Next i
    If count = 0 Then Err.Raise vbObjectError + 402, , "No TNOs for Test Name: " & name
    ReDim Preserve tnos(1 To count): ReDim Preserve volts(1 To count)
End Sub

Private Function GroupHasDUT(ByVal data As Object, ByRef tnos() As String, ByVal dut As String) As Boolean
    Dim i As Long
    For i = LBound(tnos) To UBound(tnos)
        If data.Exists(tnos(i) & "|" & dut) Then GroupHasDUT = True: Exit Function
    Next i
End Function

Private Function GroupDUTs(ByRef tnos() As String) As Collection
    Dim result As New Collection, candidates As Object, dut As Variant
    If mConfiguredDUTs.Count > 0 Then
        For Each dut In mConfiguredDUTs: result.Add CStr(dut): Next dut
    Else
        Set candidates = CreateObject("Scripting.Dictionary")
        For Each dut In mDUTs
            If GroupHasDUT(mData, tnos, CStr(dut)) Then candidates.Add CStr(dut), True
        Next dut
        Set result = SortedDUTs(candidates)
    End If
    Set GroupDUTs = result
End Function

Private Function IsNameSheet(ByVal name As String) As Boolean
    IsNameSheet = (Left$(name, 17) = "BLK_Distribution(" And Right$(name, 1) = ")")
End Function

Private Function DistributionSheetName(ByVal name As String, ByVal used As Object) As String
    Dim clean As String, bad As Variant, i As Long, suffix As String, candidate As String, sequence As Long
    clean = Trim$(name)
    For Each bad In Array("[", "]", ":", "*", "?", "/", "\")
        clean = Replace(clean, CStr(bad), "_")
    Next bad
    For i = 0 To 31: clean = Replace(clean, Chr$(i), "_"): Next i
    If Len(clean) = 0 Then Err.Raise vbObjectError + 403, , "Test Name cannot be blank."
    candidate = "BLK_Distribution(" & Left$(clean, 13) & ")"
    sequence = 1
    Do While used.Exists(candidate)
        sequence = sequence + 1: suffix = "~" & CStr(sequence)
        candidate = "BLK_Distribution(" & Left$(clean, 13 - Len(suffix)) & suffix & ")"
    Loop
    used.Add candidate, True
    DistributionSheetName = candidate
End Function

Private Sub SetupNameSheets(ByVal wb As Workbook)
    Dim used As Object, ws As Worksheet, name As Variant
    Set used = CreateObject("Scripting.Dictionary"): used.CompareMode = vbTextCompare
    For Each ws In wb.Worksheets
        If Not IsNameSheet(ws.Name) Then used.Add ws.Name, True
    Next ws
    Set mSheetNames = CreateObject("Scripting.Dictionary"): mSheetNames.CompareMode = vbTextCompare
    Set mNextRows = CreateObject("Scripting.Dictionary"): mNextRows.CompareMode = vbTextCompare
    For Each name In mGroups
        mSheetNames.Add CStr(name), DistributionSheetName(CStr(name), used)
        mNextRows.Add CStr(name), 2
    Next name
End Sub

Public Function BLK_Finalize(ByVal wb As Workbook) As String
    Dim keep As Object, name As Variant, i As Long, alerts As Boolean, text As String
    alerts = Application.DisplayAlerts
    On Error GoTo Failed
    Set keep = CreateObject("Scripting.Dictionary"): keep.CompareMode = vbTextCompare
    For Each name In mSheetNames.Items: keep.Add CStr(name), True: Next name
    Application.DisplayAlerts = False
    For i = wb.Worksheets.Count To 1 Step -1
        If IsNameSheet(wb.Worksheets(i).Name) Then
            If Not keep.Exists(wb.Worksheets(i).Name) Then wb.Worksheets(i).Delete
        End If
    Next i
    For Each name In mGroups
        ArrangeWTComparison wb.Worksheets(CStr(mSheetNames(name))), CStr(name)
    Next name
    Application.DisplayAlerts = alerts
    Exit Function
Failed:
    text = "BLK comparison layout failed: " & Err.Description
    Application.DisplayAlerts = alerts: BLK_Finalize = text
End Function

Private Sub ArrangeWTComparison(ByVal ws As Worksheet, ByVal funcName As String)
    Dim source As Variant, rows As Object, duts As Object, ordered As Collection
    Dim tnos() As String, volts() As Double, dut As Variant, xy As Variant, key As String
    Dim phase As Long, r As Long, i As Long, b As Long, last As Long, src As Long
    Dim out() As Variant, header(1 To 1, 1 To 262) As Variant, n As Long, rr As Long, dest As Long
    Dim color As Long, note As String, addVmin As Boolean
    GroupSequence mSequence, mVoltage, mTestNames, funcName, tnos, volts
    n = UBound(tnos): addVmin = HasMonitor(tnos, mMonitorTNOs)
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    source = ws.Cells(1, 1).Resize(last, 260).Value2
    Set rows = CreateObject("Scripting.Dictionary")
    Set duts = CreateObject("Scripting.Dictionary")
    For r = 2 To UBound(source, 1)
        If Left$(CStr(source(r, 1)), 2) = "WT" Then
            phase = Val(Mid$(CStr(source(r, 1)), 3)): key = ""
        ElseIf phase >= 1 And phase <= 3 Then
            If IsNumeric(source(r, 1)) And Len(CStr(source(r, 1))) > 0 Then
                key = CStr(source(r, 3)) & "|" & CStr(source(r, 4))
                If Not duts.Exists(key) Then duts.Add key, True
                rows.Add key & "|" & CStr(source(r, 1)) & "|" & phase, r
            ElseIf InStr(1, CStr(source(r, 1)), " Vmin", vbTextCompare) > 0 And Len(key) > 0 Then
                rows.Add key & "|Vmin|" & phase, r
            ElseIf Len(CStr(source(r, 1))) > 0 Then
                note = note & "WT" & phase & ": " & CStr(source(r, 1)) & vbLf
            End If
        End If
    Next r
    If mConfiguredDUTs.Count > 0 Then Set ordered = mConfiguredDUTs Else Set ordered = SortedDUTs(duts)
    If CDbl(ordered.Count) * (n + IIf(addVmin, 1, 0)) * 3 + 1 > ws.Rows.Count Then Err.Raise vbObjectError + 405, , "Comparison sheet exceeds row limit."
    header(1, 1) = "TNO / Vmin Label": header(1, 2) = "Func.": header(1, 3) = "Test Voltage"
    header(1, 4) = "X": header(1, 5) = "Y": header(1, 6) = "WT"
    For b = 0 To 255: header(1, b + 7) = "BLK " & b: Next b
    ws.UsedRange.UnMerge: ws.UsedRange.Clear
    ws.Cells(1, 1).Resize(1, 262).Value2 = header
    dest = 2
    For Each dut In ordered
        out = BuildComparisonGroup(source, rows, CStr(dut), funcName, tnos, volts, addVmin)
        rr = UBound(out, 1)
        ws.Cells(dest, 1).Resize(rr, 262).Value2 = out
        For r = 0 To rr - 1
            Select Case r Mod 3
                Case 0: color = RGB(219, 234, 254)
                Case 1: color = RGB(255, 237, 213)
                Case 2: color = RGB(220, 252, 231)
            End Select
            ws.Cells(dest + r, 1).Resize(1, 262).Interior.Color = color
        Next r
        If addVmin Then
        With ws.Cells(dest + rr - 3, 1).Resize(3, 262)
            .Interior.Color = RGB(37, 99, 168): .Font.Color = vbWhite: .Font.Bold = True
        End With
        End If
        dest = dest + rr
    Next dut
    If Len(note) > 0 Then ws.Cells(dest, 1).Value2 = Left$(note, 32000): ws.Cells(dest, 1).WrapText = True
    With ws.Cells(1, 1).Resize(dest, 262)
        .Font.Name = "Calibri": .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 26
    End With
    With ws.Cells(1, 1).Resize(1, 262)
        .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: .RowHeight = 32
    End With
    ws.Columns("A").ColumnWidth = 26: ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 16: ws.Columns("D:F").ColumnWidth = 10
    ws.Columns("G:JB").ColumnWidth = 11: ws.Columns("C").NumberFormat = "0.000"
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Cells(1, 1).Resize(IIf(dest > 2, dest - 1, 1), 262).AutoFilter
End Sub

Private Function BuildComparisonGroup(ByRef source As Variant, ByVal rows As Object, ByVal dut As String, ByVal funcName As String, ByRef tnos() As String, ByRef volts() As Double, Optional ByVal addVmin As Boolean = True) As Variant
    Dim out() As Variant, xy As Variant, n As Long, rr As Long, i As Long, phase As Long, key As String, src As Long, b As Long
    n = UBound(tnos)
        xy = Split(dut, "|")
        ReDim out(1 To (n + IIf(addVmin, 1, 0)) * 3, 1 To 262)
        rr = 0
        For i = 1 To n + IIf(addVmin, 1, 0)
            For phase = 1 To 3
                rr = rr + 1
                If i <= n Then
                    out(rr, 1) = tnos(i): out(rr, 3) = volts(i)
                    key = dut & "|" & tnos(i) & "|" & phase
                Else
                    out(rr, 1) = "X" & xy(0) & "Y" & xy(1) & " Vmin"
                    key = dut & "|Vmin|" & phase
                End If
                out(rr, 2) = funcName: out(rr, 4) = CLng(xy(0)): out(rr, 5) = CLng(xy(1)): out(rr, 6) = "WT" & phase
                src = 0: If rows.Exists(key) Then src = CLng(rows(key))
                For b = 0 To 255
                    If src > 0 Then
                        out(rr, b + 7) = source(src, b + 5)
                    ElseIf i <= n Then
                        out(rr, b + 7) = "Not Found"
                    Else
                        out(rr, b + 7) = "No Data"
                    End If
                Next b
            Next phase
        Next i
    BuildComparisonGroup = out
End Function

Public Function BLK_AddressTNOs() As Object
    Dim result As Object, i As Long
    Set result = CreateObject("Scripting.Dictionary")
    For i = LBound(mSequence) To UBound(mSequence)
        result.Add mSequence(i), result.Count
    Next i
    Set BLK_AddressTNOs = result
End Function

Private Sub ReadItemColumns(ByVal data As Variant, ByVal monitor As Boolean, ByVal seen As Object, ByRef tnos() As String, ByRef names() As String, ByRef volts() As Double, ByRef count As Long)
    Dim i As Long, tno As String, label As String, j As Long
    label = IIf(monitor, "Monitor", "Function")
    For i = LBound(data, 1) To UBound(data, 1)
        For j = 1 To 3
            If IsError(data(i, j)) Then Err.Raise vbObjectError + 302, , label & " cell error at row " & i + 1
        Next j
        If Len(Trim$(CStr(data(i, 1)))) > 0 Then
            tno = Trim$(CStr(data(i, 1)))
            If Len(tno) <> 3 Then Err.Raise vbObjectError + 302, , label & " TNO must have 3 digits at row " & i + 1
            tno = CStr(CLng(UnsignedValue(tno, 999)))
            If CLng(tno) < 100 Then Err.Raise vbObjectError + 302, , label & " TNO must be 100-999."
            If seen.Exists(tno) Then Err.Raise vbObjectError + 303, , "Duplicate TNO across Monitor/Function: " & tno
            If Len(Trim$(CStr(data(i, 2)))) = 0 Then Err.Raise vbObjectError + 401, , label & " Func. missing at row " & i + 1
            If Len(Trim$(CStr(data(i, 3)))) = 0 Or Not IsNumeric(data(i, 3)) Or VarType(data(i, 3)) = vbBoolean Then Err.Raise vbObjectError + 304, , label & " voltage missing/invalid at row " & i + 1
            count = count + 1
            tnos(count) = tno: names(count) = Trim$(CStr(data(i, 2))): volts(count) = CDbl(data(i, 3))
            seen.Add tno, True
            If monitor Then mMonitorTNOs.Add tno, True
        ElseIf Len(Trim$(CStr(data(i, 2)))) > 0 Or Len(Trim$(CStr(data(i, 3)))) > 0 Then
            Err.Raise vbObjectError + 304, , label & " Func./voltage without TNO at row " & i + 1
        End If
    Next i
End Sub

Private Function HasMonitor(ByRef tnos() As String, ByVal monitor As Object) As Boolean
    Dim i As Long
    For i = LBound(tnos) To UBound(tnos)
        If monitor.Exists(tnos(i)) Then HasMonitor = True: Exit Function
    Next i
End Function

Private Function BuildItemDUTGroup(ByVal data As Object, ByRef tnos() As String, ByRef volts() As Double, ByVal dut As String, ByVal threshold As Double, ByVal monitor As Object) As Variant
    Dim output() As Variant, onlyTNOs() As String, onlyVolts() As Double, judged As Variant
    Dim values As Variant, xy As Variant, key As String, i As Long, b As Long, n As Long, count As Long
    n = UBound(tnos): xy = Split(dut, "|")
    ReDim onlyTNOs(1 To n): ReDim onlyVolts(1 To n)
    For i = LBound(tnos) To UBound(tnos)
        If monitor.Exists(tnos(i)) Then
            count = count + 1: onlyTNOs(count) = tnos(i): onlyVolts(count) = volts(i)
        End If
    Next i
    ReDim output(1 To n + IIf(count > 0, 1, 0), 1 To 260)
    For i = LBound(tnos) To UBound(tnos)
        output(i, 1) = tnos(i): output(i, 2) = volts(i): output(i, 3) = CLng(xy(0)): output(i, 4) = CLng(xy(1))
        key = tnos(i) & "|" & dut
        If data.Exists(key) Then
            values = data(key)
            For b = 0 To 255: output(i, b + 5) = values(b): Next b
        Else
            For b = 0 To 255: output(i, b + 5) = "Not Found": Next b
        End If
    Next i
    If count > 0 Then
        ReDim Preserve onlyTNOs(1 To count): ReDim Preserve onlyVolts(1 To count)
        judged = BuildDUTGroup(data, onlyTNOs, onlyVolts, dut, threshold)
        For b = 1 To 260: output(n + 1, b) = judged(count + 1, b): Next b
    End If
    BuildItemDUTGroup = output
End Function
