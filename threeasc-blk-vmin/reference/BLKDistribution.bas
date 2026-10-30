Attribute VB_Name = "ImportFailAddresses"
Option Explicit
Private Const READ_BYTES As Long = 262144
Private Const MAX_EXACT As Double = 9007199254740991#
Private Const MAX_LINE As Long = 1048576
Private mTags As Object, mTagHeads As Object, mBlockTag As Object, mSpaces As Object, mDigits As Object
Private mData As Object, mWanted As Object, mDUTs As Object
Private mTno As String, mX As String, mY As String, mPath As String
Private mActive As Boolean, mRecordKey As String, mRecordDUT As String
Private mValues() As Double, mCount As Long, mLine As Long, mRecordLine As Long

Public Sub ImportASCFailAddresses()
    RunBLKDistribution
End Sub

Public Sub RunBLKDistribution()
    Static busy As Boolean
    Dim ws As Worksheet, fd As Object, paths As Collection, path As Variant
    Dim tnos() As String, volts() As Double, dutList As Collection, threshold As Double
    Dim screen As Boolean, events As Boolean, alerts As Boolean, calc As XlCalculation, status As Variant
    Dim message As String, result As String, ok As Boolean, success As Long, failed As Long, started As Date
    If busy Then Exit Sub
    busy = True
    screen = Application.ScreenUpdating: events = Application.EnableEvents
    alerts = Application.DisplayAlerts: calc = Application.Calculation: status = Application.StatusBar
    On Error GoTo Failed
    Set ws = ThisWorkbook.Worksheets("DataPath")
    ReadConfiguration ws, tnos, volts, dutList, threshold
    Set fd = Application.FileDialog(3): Set paths = New Collection
    With fd
        .Title = "Select ASC files for BLK Distribution": .AllowMultiSelect = True
        .Filters.Clear: .Filters.Add "ASC raw data", "*.asc"
        If .Show <> -1 Then GoTo Cleanup
        For Each path In .SelectedItems: paths.Add CStr(path): Next path
    End With
    started = Now
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.DisplayAlerts = False: Application.Calculation = xlCalculationManual
    For Each path In paths
        result = ProcessFile(CStr(path), tnos, volts, dutList, threshold, ok)
        If ok Then success = success + 1 Else failed = failed + 1
        If paths.Count <= 10 Or Not ok Then message = message & result & vbCrLf
    Next path
    message = "Completed: " & success & "; failed: " & failed & vbCrLf & message & "Run time: " & DateDiff("s", started, Now) & " seconds"
Cleanup:
    On Error Resume Next
    Set mData = Nothing: Set mWanted = Nothing: Set mDUTs = Nothing
    Application.Calculation = calc: Application.DisplayAlerts = alerts
    Application.EnableEvents = events: Application.ScreenUpdating = screen: Application.StatusBar = status
    busy = False
    If Len(message) > 0 Then MsgBox message, vbInformation
    Exit Sub
Failed:
    message = "BLK Distribution failed (" & Err.Number & "): " & Err.Description
    Resume Cleanup
End Sub

Private Sub ReadConfiguration(ByVal ws As Worksheet, ByRef tnos() As String, ByRef volts() As Double, ByRef dutList As Collection, ByRef threshold As Double)
    Dim data As Variant, xy As Variant, last As Long, i As Long, count As Long, tno As String, key As String
    Dim seen As Object, duts As Object, value As Variant
    Set seen = CreateObject("Scripting.Dictionary"): Set duts = CreateObject("Scripting.Dictionary")
    Set dutList = New Collection
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last < 2 Then Err.Raise vbObjectError + 301, , "Enter TNO and voltage in DataPath!A2:B..."
    data = ws.Range("A2:B" & last).Value2
    ReDim tnos(1 To UBound(data, 1)): ReDim volts(1 To UBound(data, 1))
    For i = 1 To UBound(data, 1)
        If IsError(data(i, 1)) Or IsError(data(i, 2)) Then Err.Raise vbObjectError + 302, , "Configuration cell error at row " & i + 1
        If Len(Trim$(CStr(data(i, 1)))) > 0 Then
            tno = CStr(CLng(UnsignedValue(CStr(data(i, 1)), 2147483647#)))
            If tno = "0" Then Err.Raise vbObjectError + 302, , "TNO must be a positive integer."
            If seen.Exists(tno) Then Err.Raise vbObjectError + 303, , "Duplicate configured TNO: " & tno
            If Len(Trim$(CStr(data(i, 2)))) = 0 Or Not IsNumeric(data(i, 2)) Or VarType(data(i, 2)) = vbBoolean Then Err.Raise vbObjectError + 304, , "Missing/invalid voltage at DataPath row " & i + 1
            count = count + 1: tnos(count) = tno: volts(count) = CDbl(data(i, 2)): seen.Add tno, True
        ElseIf Len(Trim$(CStr(data(i, 2)))) > 0 Then
            Err.Raise vbObjectError + 304, , "Voltage without TNO at DataPath row " & i + 1
        End If
    Next i
    If count = 0 Then Err.Raise vbObjectError + 301, , "No configured TNO."
    ReDim Preserve tnos(1 To count): ReDim Preserve volts(1 To count)
    value = ws.Range("D2").Value2
    If IsError(value) Then Err.Raise vbObjectError + 305, , "Invalid VminFBCThreshold."
    threshold = 0
    If Len(Trim$(CStr(value))) > 0 Then threshold = UnsignedValue(CStr(value), MAX_EXACT)
    last = ws.Cells(ws.Rows.Count, 6).End(xlUp).Row
    i = ws.Cells(ws.Rows.Count, 7).End(xlUp).Row: If i > last Then last = i
    If last >= 2 Then
        xy = ws.Range("F2:G" & last).Value2
        For i = 1 To UBound(xy, 1)
            If IsError(xy(i, 1)) Or IsError(xy(i, 2)) Then Err.Raise vbObjectError + 306, , "Invalid DUT list at row " & i + 1
            If Len(Trim$(CStr(xy(i, 1)))) > 0 Or Len(Trim$(CStr(xy(i, 2)))) > 0 Then
                key = Coordinate(CStr(xy(i, 1))) & "|" & Coordinate(CStr(xy(i, 2)))
                If Not duts.Exists(key) Then dutList.Add key: duts.Add key, True
            End If
        Next i
    End If
End Sub

Private Function ProcessFile(ByVal path As String, ByRef tnos() As String, ByRef volts() As Double, ByVal configuredDUTs As Collection, ByVal threshold As Double, ByRef ok As Boolean) As String
    Dim report As Workbook, ws As Worksheet, fso As Object, dutList As Collection, key As Variant
    Dim i As Long, block As Long, firstRow As Long, groupSize As Long, totalRows As Double
    Dim header(1 To 1, 1 To 260) As Variant, group As Variant, outputPath As String, tempPath As String
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    ok = False: Set fso = CreateObject("Scripting.FileSystemObject")
    Set mWanted = CreateObject("Scripting.Dictionary")
    For i = LBound(tnos) To UBound(tnos): mWanted.Add tnos(i), True: Next i
    ParseASCFile path
    Set dutList = New Collection
    If configuredDUTs.Count > 0 Then
        For Each key In configuredDUTs: dutList.Add CStr(key): Next key
    Else
        Set dutList = SortedDUTs(mDUTs)
    End If
    If dutList.Count = 0 Then Err.Raise vbObjectError + 307, , "No valid FBC records for the configured TNOs."
    groupSize = UBound(tnos) - LBound(tnos) + 2
    totalRows = CDbl(dutList.Count) * groupSize + 1
    If totalRows > 1048576# Then Err.Raise vbObjectError + 308, , "BLK_Distribution exceeds Excel row limit."
    Set report = Workbooks.Add(xlWBATWorksheet)
    EnsureModernReport report, tempPath
    Set ws = report.Worksheets(1): ws.Name = "BLK_Distribution"
    header(1, 1) = "TNO / Vmin Label": header(1, 2) = "Test Voltage": header(1, 3) = "DUT X": header(1, 4) = "DUT Y"
    For block = 0 To 255: header(1, block + 5) = "BLK " & block: Next block
    ws.Range("A1:IZ1").Value2 = header
    firstRow = 2
    For Each key In dutList
        group = BuildDUTGroup(mData, tnos, volts, CStr(key), threshold)
        WriteGroup ws, firstRow, group
        firstRow = firstRow + groupSize
        Application.StatusBar = "BLK Distribution: " & fso.GetFileName(path) & "; rows=" & firstRow - 1
        DoEvents
    Next key
    With ws.Range("A1:IZ" & CLng(totalRows))
        .Font.Name = "Calibri": .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 26
    End With
    ws.Columns("A").ColumnWidth = 26: ws.Columns("B").ColumnWidth = 15
    ws.Columns("C:D").ColumnWidth = 10: ws.Columns("E:IZ").ColumnWidth = 11
    ws.Columns("B").NumberFormat = "0.000"
    With ws.Range("A1:IZ1")
        .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True
        .RowHeight = 32
    End With
    firstRow = 2 + groupSize - 1
    For i = 1 To dutList.Count
        ws.Cells(firstRow, 1).Resize(1, 4).Merge
        With ws.Cells(firstRow, 1).Resize(1, 260)
            .Interior.Color = RGB(37, 99, 168): .Font.Color = vbWhite: .Font.Bold = True
        End With
        firstRow = firstRow + groupSize
    Next i
    outputPath = UniquePath(fso, fso.GetParentFolderName(path), fso.GetBaseName(path) & "_BLK_Distribution")
    report.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    report.Close SaveChanges:=False: Set report = Nothing
    If Len(tempPath) > 0 Then If fso.FileExists(tempPath) Then fso.DeleteFile tempPath, True
    ok = True: ProcessFile = "Saved: " & outputPath
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If Not report Is Nothing Then report.Close SaveChanges:=False
    If Len(tempPath) > 0 Then If fso.FileExists(tempPath) Then fso.DeleteFile tempPath, True
    ProcessFile = "FAILED: " & path & vbCrLf & "(" & errorNumber & ") " & errorText
End Function

Private Sub ParseASCFile(ByVal path As String)
    Dim no As Integer, opened As Boolean, encoding As String, buffer As String, raw As String
    Dim p As Long, fso As Object, errorNumber As Long, errorText As String
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.GetFile(path).Size > 2147483647# Then Err.Raise vbObjectError + 309, , "ASC exceeds 2 GB reader limit."
    ResetParser path
    no = FreeFile: Open path For Binary Access Read As #no: opened = True
    DetectEncoding no, encoding
    Do While Seek(no) <= LOF(no)
        raw = ReadASCChunk(no, encoding): buffer = buffer & raw
        Do
            p = InStr(buffer, vbLf)
            If p = 0 Then Exit Do
            mLine = mLine + 1: ParseLine Replace(Left$(buffer, p - 1), vbCr, "")
            buffer = Mid$(buffer, p + 1)
        Loop
        If Len(buffer) > MAX_LINE Then Err.Raise vbObjectError + 310, , "ASC line exceeds 1 MB."
    Loop
    If Len(buffer) > 0 Then mLine = mLine + 1: ParseLine Replace(buffer, vbCr, "")
    FinishRecord
    Close #no: opened = False
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    If opened Then Close #no
    On Error GoTo 0
    Err.Raise errorNumber, , errorText & " [file=" & path & "; line=" & mLine & "]"
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

Private Function UniquePath(ByVal fso As Object, ByVal folder As String, ByVal base As String) As String
    Dim result As String, sequence As Long
    result = fso.BuildPath(folder, base & ".xlsx")
    Do While fso.FileExists(result)
        sequence = sequence + 1: result = fso.BuildPath(folder, base & "_" & Format$(sequence, "000") & ".xlsx")
    Loop
    If Len(result) > 218 Then Err.Raise vbObjectError + 317, , "Output path exceeds Excel 2016 limit."
    UniquePath = result
End Function

' Existing validated binary reader and modern-workbook guard follow.
Private Sub DetectEncoding(ByVal fileNo As Integer, ByRef encoding As String)
    Dim bytes() As Byte, count As Long
    encoding = "system ANSI"
    count = LOF(fileNo): If count > 3 Then count = 3
    If count = 0 Then Exit Sub
    ReDim bytes(0 To count - 1): Get #fileNo, 1, bytes
    Seek #fileNo, 1
    If count >= 2 Then
        If bytes(0) = 255 And bytes(1) = 254 Then encoding = "unicode": Seek #fileNo, 3: Exit Sub
        If bytes(0) = 254 And bytes(1) = 255 Then encoding = "unicodeFFFE": Seek #fileNo, 3: Exit Sub
    End If
    If count = 3 Then
        If bytes(0) = 239 And bytes(1) = 187 And bytes(2) = 191 Then encoding = "utf-8": Seek #fileNo, 4
    End If
End Sub

Private Function ReadASCChunk(ByVal fileNo As Integer, ByVal encoding As String) As String
    Dim bytes() As Byte, count As Long, lead As Long, expected As Long, tail As Long
    Dim stream As Object
    count = LOF(fileNo) - Seek(fileNo) + 1
    If count > READ_BYTES Then count = READ_BYTES
    If count <= 0 Then Exit Function
    ReDim bytes(0 To count - 1)
    Get #fileNo, , bytes
    If encoding = "utf-8" And Seek(fileNo) <= LOF(fileNo) Then
        lead = count - 1
        Do While lead > 0
            If (bytes(lead) And &HC0) <> &H80 Then Exit Do
            lead = lead - 1
        Loop
        expected = 1
        If bytes(lead) >= &HC2 And bytes(lead) <= &HDF Then expected = 2
        If bytes(lead) >= &HE0 And bytes(lead) <= &HEF Then expected = 3
        If bytes(lead) >= &HF0 And bytes(lead) <= &HF4 Then expected = 4
        tail = count - lead
        If tail < expected Then
            Seek #fileNo, Seek(fileNo) - tail
            ReDim Preserve bytes(0 To lead - 1)
        End If
    End If
    If encoding = "system ANSI" Then
        ' Retain the trailing non-ASCII byte run so a DBCS character is not split.
        If Seek(fileNo) <= LOF(fileNo) Then
            lead = UBound(bytes)
            Do While lead >= 0
                If bytes(lead) < 128 Then Exit Do
                lead = lead - 1
            Loop
            tail = UBound(bytes) - lead
            If tail > 0 And lead >= 0 Then
                Seek #fileNo, Seek(fileNo) - tail
                ReDim Preserve bytes(0 To lead)
            End If
        End If
        ReadASCChunk = StrConv(bytes, vbUnicode)
    Else
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1: stream.Open: stream.Write bytes
        stream.Position = 0: stream.Type = 2: stream.Charset = encoding
        ReadASCChunk = stream.ReadText
        stream.Close
    End If
End Function


Private Sub EnsureModernReport(ByRef report As Workbook, ByRef tempPath As String)
    Dim fso As Object
    If report.Worksheets(1).Rows.Count >= 1048576 And report.Worksheets(1).Columns.Count >= 16384 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempPath = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName & ".xlsx")
    report.SaveAs Filename:=tempPath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    If report.Worksheets(1).Rows.Count < 1048576 Or report.Worksheets(1).Columns.Count < 16384 Then
        report.Close SaveChanges:=False
        Set report = Workbooks.Open(Filename:=tempPath, UpdateLinks:=0, ReadOnly:=False, AddToMru:=False)
    End If
    If report.Worksheets(1).Rows.Count < 1048576 Or report.Worksheets(1).Columns.Count < 16384 Then
        Err.Raise vbObjectError + 207, , "Report remains in legacy Excel compatibility mode. An XLSX report with 1048576 rows is required."
    End If
End Sub

