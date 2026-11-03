Attribute VB_Name = "BLKVminEngine"
Option Explicit
Option Private Module
Private Const READ_BYTES As Long = 262144
Private Const MAX_EXACT As Double = 9007199254740991#
Private Const MAX_LINE As Long = 1048576
Private mTags As Object, mBlockTag As Object, mSpaces As Object, mDigits As Object
Private mData As Object, mWanted As Object, mDUTs As Object
Private mTno As String, mX As String, mY As String, mPath As String
Private mActive As Boolean, mRecordKey As String, mRecordDUT As String
Private mValues() As Double, mCount As Long, mLine As Long, mRecordLine As Long

Private mSequence() As String, mCarry As String, mBLKError As String, mNextRow As Long
Private mIgnored As Long
Private mLocations As Object, mPhaseErrors As Object, mLocationCarry As String, mLocationError As String

Public Sub BLK_Setup(ByVal cfg As Worksheet)
    Dim cols As Variant, col As Variant, data As Variant, i As Long, count As Long, tno As String, seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    ReDim mSequence(1 To 900)
    cols = Array("A2:A901", "D2:D901")
    For Each col In cols
        data = cfg.Range(CStr(col)).Value2
        For i = LBound(data, 1) To UBound(data, 1)
            If IsError(data(i, 1)) Then Err.Raise vbObjectError + 302, , "Invalid TNO in Config " & col & "; row=" & i + 1
            tno = Trim$(CStr(data(i, 1)))
            If Len(tno) > 0 Then
                If Len(tno) <> 3 Then Err.Raise vbObjectError + 302, , "TNO must be 100-999; row=" & i + 1
                tno = CStr(CLng(UnsignedValue(tno, 999)))
                If CLng(tno) < 100 Then Err.Raise vbObjectError + 302, , "TNO must be 100-999; row=" & i + 1
                If seen.Exists(tno) Then Err.Raise vbObjectError + 303, , "Duplicate TNO: " & tno
                count = count + 1: mSequence(count) = tno: seen.Add tno, True
            End If
        Next i
    Next col
    If count = 0 Then Err.Raise vbObjectError + 301, , "Enter TNOs in Config A2:A901 or D2:D901."
    ReDim Preserve mSequence(1 To count)
    mNextRow = 2
    Set mLocations = CreateObject("Scripting.Dictionary")
    Set mPhaseErrors = CreateObject("Scripting.Dictionary")
    mLocationCarry = "": mLocationError = ""
End Sub

Public Function BLK_AddressTNOs() As Object
    Dim result As Object, i As Long
    Set result = CreateObject("Scripting.Dictionary")
    For i = LBound(mSequence) To UBound(mSequence): result.Add mSequence(i), result.Count: Next i
    Set BLK_AddressTNOs = result
End Function

Public Function BLK_WritePhase(ByVal wb As Workbook, ByVal phase As Long) As String
    Dim ws As Worksheet, header(1 To 1, 1 To 260) As Variant, packet() As Variant
    Dim duts As Collection, dut As Variant, xy As Variant, values As Variant, key As String
    Dim i As Long, b As Long, count As Long, take As Long, remaining As Long, color As Long, startRow As Long
    On Error GoTo Failed
    mPhaseErrors("WT" & phase) = mBLKError
    Set ws = wb.Worksheets("FBC by BLK")
    If phase = 1 Then
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
        ws.UsedRange.UnMerge: ws.UsedRange.Clear
        header(1, 1) = "WT": header(1, 2) = "TNO": header(1, 3) = "X": header(1, 4) = "Y"
        For b = 0 To 255: header(1, b + 5) = "BLK " & b: Next b
        ws.Cells(1, 1).Resize(1, 260).Value2 = header
        With ws.Cells(1, 1).Resize(1, 260)
            .Font.Name = "Calibri": .Font.Size = 11: .Font.Bold = True: .Font.Color = vbWhite
            .Interior.Color = RGB(23, 50, 77): .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 32
        End With
        ws.Columns("A:D").ColumnWidth = 10: ws.Columns("E:IZ").ColumnWidth = 12
        mNextRow = 2
    End If
    If Len(mBLKError) > 0 Then BLK_WritePhase = mBLKError: Exit Function
    If mData.Count = 0 Then BLK_WritePhase = "No selected FBC records; ignored unselected/zero-TNO FBC records=" & mIgnored: Exit Function
    If CDbl(mNextRow) + mData.Count - 1 > ws.Rows.Count Then Err.Raise vbObjectError + 404, , "FBC by BLK exceeds Excel row limit."
    Set duts = SortedDUTs(mDUTs): startRow = mNextRow: remaining = mData.Count
    take = 512: If remaining < take Then take = remaining
    ReDim packet(1 To take, 1 To 260)
    For Each dut In duts
        xy = Split(CStr(dut), "|")
        For i = LBound(mSequence) To UBound(mSequence)
            key = mSequence(i) & "|" & CStr(dut)
            If mData.Exists(key) Then
                count = count + 1: values = mData(key)
                packet(count, 1) = "WT" & phase: packet(count, 2) = CLng(mSequence(i))
                packet(count, 3) = CLng(xy(0)): packet(count, 4) = CLng(xy(1))
                For b = 0 To 255: packet(count, b + 5) = values(b): Next b
                If count = take Then
                    ws.Cells(mNextRow, 1).Resize(take, 260).Value2 = packet
                    mNextRow = mNextRow + take: remaining = remaining - take: count = 0
                    If remaining > 0 Then
                        take = 512: If remaining < take Then take = remaining
                        ReDim packet(1 To take, 1 To 260)
                    End If
                End If
            End If
        Next i
    Next dut
    Select Case phase
        Case 1: color = RGB(219, 234, 254)
        Case 2: color = RGB(255, 237, 213)
        Case 3: color = RGB(220, 252, 231)
    End Select
    With ws.Cells(startRow, 1).Resize(mNextRow - startRow, 260)
        .Font.Name = "Calibri": .Font.Size = 11: .NumberFormat = "0"
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 26: .Interior.Color = color
    End With
    Exit Function
Failed:
    BLK_WritePhase = "FBC output error (" & Err.Number & "): " & Err.Description
    mPhaseErrors("WT" & phase) = BLK_WritePhase
End Function

Public Function BLK_Finalize(ByVal wb As Workbook) As String
    Dim ws As Worksheet, ranks As Object, values As Variant, helper As Variant, i As Long, lastRow As Long
    On Error GoTo Failed
    Set ws = wb.Worksheets("FBC by BLK")
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    lastRow = mNextRow - 1
    If lastRow > 1 Then
        Set ranks = CreateObject("Scripting.Dictionary")
        For i = LBound(mSequence) To UBound(mSequence): ranks.Add mSequence(i), i: Next i
        ' Read only the TNO column, not the 256-column FBC table.
        values = ws.Cells(2, 2).Resize(lastRow - 1, 1).Value2
        helper = BuildTNORanks(values, ranks, lastRow - 1)
        ws.Cells(1, 261).Value2 = "Temporary TNO order"
        ws.Cells(2, 261).Resize(lastRow - 1, 1).Value2 = helper
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Cells(2, 3).Resize(lastRow - 1, 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=ws.Cells(2, 4).Resize(lastRow - 1, 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=ws.Cells(2, 261).Resize(lastRow - 1, 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SortFields.Add Key:=ws.Cells(2, 1).Resize(lastRow - 1, 1), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange ws.Cells(1, 1).Resize(lastRow, 261)
            .Header = xlYes: .MatchCase = False: .Orientation = xlTopToBottom
            .Apply
            .SortFields.Clear
        End With
        ws.Cells(1, 261).Resize(lastRow, 1).ClearContents
    End If
    ws.Cells(1, 1).Resize(lastRow, 260).AutoFilter
    WriteMapLocations wb
    Exit Function
Failed:
    BLK_Finalize = "FBC final formatting: " & Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Sort.SortFields.Clear
        If lastRow > 0 Then ws.Cells(1, 261).Resize(lastRow, 1).ClearContents
    End If
End Function

Private Function BuildTNORanks(ByVal values As Variant, ByVal ranks As Object, ByVal rowCount As Long) As Variant
    Dim result() As Variant, i As Long, value As Variant, key As String
    ReDim result(1 To rowCount, 1 To 1)
    For i = 1 To rowCount
        If rowCount = 1 Then value = values Else value = values(i, 1)
        key = CStr(CLng(value))
        If Not ranks.Exists(key) Then Err.Raise vbObjectError + 405, , "Output TNO is not configured: " & key
        result(i, 1) = CLng(ranks(key))
    Next i
    BuildTNORanks = result
End Function

Public Sub BLK_Release()
    Set mData = Nothing: Set mWanted = Nothing: Set mDUTs = Nothing
    Set mLocations = Nothing: Set mPhaseErrors = Nothing
    mCarry = "": mActive = False
End Sub

Public Sub BLK_BeginFile(ByVal path As String)
    Dim i As Long
    Set mWanted = CreateObject("Scripting.Dictionary")
    For i = LBound(mSequence) To UBound(mSequence): mWanted.Add mSequence(i), True: Next i
    ResetParser path
    mCarry = "": mBLKError = "": mIgnored = 0: mLocationCarry = ""
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
        Call FinishRecord
        mCarry = ""
    End If
    Exit Sub
Failed:
    mBLKError = "FBC Parse Error (" & Err.Number & "): " & Err.Description & "; line=" & mLine
    mCarry = "": mActive = False
End Sub


Private Sub ResetParser(ByVal path As String)
    Set mData = CreateObject("Scripting.Dictionary"): Set mDUTs = CreateObject("Scripting.Dictionary")
    mTno = "": mX = "": mY = "": mPath = path: mLine = 0
    mActive = False: mCount = 0
    If mTags Is Nothing Then
        Set mTags = CreateObject("VBScript.RegExp"): mTags.Global = True: mTags.IgnoreCase = True
        mTags.Pattern = "\b(Tno|DUT|X|Y)\s*=\s*([+-]?[0-9]+)(?=\s|[,;]|$)"
        Set mBlockTag = CreateObject("VBScript.RegExp"): mBlockTag.IgnoreCase = True
        mBlockTag.Pattern = "\bFBC\s+BY\s+BLK\s*=\s*"
        Set mSpaces = CreateObject("VBScript.RegExp"): mSpaces.Global = True: mSpaces.Pattern = "[\s,;]+"
        Set mDigits = CreateObject("VBScript.RegExp"): mDigits.Pattern = "^[0-9]+$"
    End If
End Sub

Private Sub ParseLine(ByVal line As String)
    Dim blocks As Object, tags As Object, before As String, payload As String, text As String
    Dim currentTNO As String, chipX As String, chipY As String
    line = NormalizeASCWhitespace(line)
    Set blocks = mBlockTag.Execute(line)
    If blocks.Count = 0 Then
        ' Ordinary metadata never establishes or changes a Chip/TNO context.
        If Not mActive Then Exit Sub
        text = Trim$(line)
        If Len(text) = 0 Then Exit Sub
        If mCount = 256 And InStr(1, "0123456789,+-", Left$(text, 1), vbBinaryCompare) = 0 Then
            FinishRecord
        Else
            AddValues line
        End If
        Exit Sub
    End If
    FinishRecord
    before = Left$(line, blocks(0).FirstIndex)
    Set tags = mTags.Execute(before)
    currentTNO = RecordHeaderInteger(tags, "TNO")
    If Not mWanted.Exists(currentTNO) Then mIgnored = mIgnored + 1: Exit Sub
    chipX = RecordHeaderInteger(tags, "X")
    chipY = RecordHeaderInteger(tags, "Y")
    mRecordDUT = chipX & "|" & chipY
    mRecordKey = currentTNO & "|" & mRecordDUT
    mRecordLine = mLine: mActive = True: mCount = 0: ReDim mValues(0 To 255)
    payload = Mid$(line, blocks(0).FirstIndex + blocks(0).Length + 1)
    AddValues payload
End Sub

Private Function RecordHeaderInteger(ByVal tags As Object, ByVal field As String) As String
    Dim tag As Object, count As Long, value As String
    For Each tag In tags
        If UCase$(tag.SubMatches(0)) = field Then
            count = count + 1: value = CStr(tag.SubMatches(1))
        End If
    Next tag
    If count = 0 Then Err.Raise vbObjectError + 318, , "FBC record requires integer " & field & " on the same header line; record line=" & mLine
    If count > 1 Then Err.Raise vbObjectError + 318, , "FBC record has multiple " & field & " fields; record line=" & mLine
    RecordHeaderInteger = Coordinate(value)
End Function

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


Private Function SortedDUTs(ByVal duts As Object) As Collection
    Dim result As New Collection, keys As Variant, key As Variant
    If duts.Count > 0 Then
        keys = duts.Keys: SortXY keys, LBound(keys), UBound(keys)
        For Each key In keys: result.Add CStr(key): Next key
    End If
    Set SortedDUTs = result
End Function

Private Function NormalizeASCWhitespace(ByVal text As String) As String
    Dim code As Long
    text = Replace(text, ChrW(&HA0), " ")
    text = Replace(text, ChrW(&H3000), " ")
    For code = &H2000 To &H200A
        text = Replace(text, ChrW(code), " ")
    Next code
    text = Replace(text, ChrW(&H202F), " ")
    text = Replace(text, ChrW(&H205F), " ")
    If Left$(text, 1) = ChrW(&HFEFF) Then text = Mid$(text, 2)
    NormalizeASCWhitespace = text
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


Public Function BLK_RecordCount() As Long
    If Len(mBLKError) = 0 Then BLK_RecordCount = mData.Count
End Function

Public Function BLK_Diagnostics() As String
    BLK_Diagnostics = "line=" & mLine & "; active=" & mActive & "; values=" & mCount & "; carry=" & Len(mCarry) & "; ignored=" & mIgnored & "; error=" & mBLKError
End Function

Private Function CountFailBlocks(ByVal values As Variant) As Long
    Dim block As Long, count As Long
    For block = LBound(values) To UBound(values)
        If CDbl(values(block)) > 0 Then count = count + 1
    Next block
    CountFailBlocks = count
End Function

Public Sub BLK_ObserveLocations(ByVal raw As String, ByVal atEOF As Boolean)
    Dim p As Long, line As String
    On Error GoTo Failed
    If mLocations Is Nothing Then Set mLocations = CreateObject("Scripting.Dictionary")
    mLocationCarry = mLocationCarry & raw
    Do
        p = InStr(mLocationCarry, vbLf): If p = 0 Then Exit Do
        line = Replace(Left$(mLocationCarry, p - 1), vbCr, "")
        CaptureMapLocation line
        mLocationCarry = Mid$(mLocationCarry, p + 1)
    Loop
    If Len(mLocationCarry) > MAX_LINE Then Err.Raise vbObjectError + 450, , "Location line exceeds 1 MB."
    If atEOF Then CaptureMapLocation mLocationCarry: mLocationCarry = ""
    Exit Sub
Failed:
    mLocationError = "Location capture error: " & Err.Description
    mLocationCarry = ""
End Sub

Private Sub CaptureMapLocation(ByVal line As String)
    Dim tags As Object, tag As Object, x As String, y As String, nx As Long, ny As Long, key As String, blocks As Object, isDUTRecord As Boolean
    If InStr(line, "=") = 0 Then Exit Sub
    line = NormalizeASCWhitespace(line)
    Set blocks = mBlockTag.Execute(line)
    If blocks.Count > 0 Then line = Left$(line, blocks(0).FirstIndex)
    Set tags = mTags.Execute(line)
    For Each tag In tags
        Select Case UCase$(tag.SubMatches(0))
            Case "TNO", "DUT": isDUTRecord = True
            Case "X": nx = nx + 1: x = Coordinate(CStr(tag.SubMatches(1)))
            Case "Y": ny = ny + 1: y = Coordinate(CStr(tag.SubMatches(1)))
        End Select
    Next tag
    If nx = 1 And ny = 1 And isDUTRecord Then
        key = x & "|" & y
        If Not mLocations.Exists(key) Then mLocations.Add key, True
    End If
End Sub

Private Sub WriteMapLocations(ByVal wb As Workbook)
    Dim ws As Worksheet, ordered As Collection, item As Variant, xy As Variant, data() As Variant, i As Long
    Set ws = wb.Worksheets("Chip Locations")
    ws.UsedRange.Clear
    ws.Range("A1:B1").Value2 = Array("X", "Y")
    ws.Range("D1:F1").Value2 = Array("WT", "FBC Status", "Message")
    Set ordered = SortedDUTs(mLocations)
    If ordered.Count > 0 Then
        ReDim data(1 To ordered.Count, 1 To 2)
        For Each item In ordered
            i = i + 1: xy = Split(CStr(item), "|"): data(i, 1) = CLng(xy(0)): data(i, 2) = CLng(xy(1))
        Next item
        ws.Range("A2").Resize(ordered.Count, 2).Value2 = data
    End If
    For i = 1 To 3
        ws.Cells(i + 1, 4).Value2 = "WT" & i
        If mPhaseErrors.Exists("WT" & i) Then
            If Len(CStr(mPhaseErrors("WT" & i))) > 0 Then ws.Cells(i + 1, 5).Value2 = "Data Error" Else ws.Cells(i + 1, 5).Value2 = "OK"
            ws.Cells(i + 1, 6).Value2 = mPhaseErrors("WT" & i)
        Else
            ws.Cells(i + 1, 5).Value2 = "Data Error": ws.Cells(i + 1, 6).Value2 = "WT not processed."
        End If
    Next i
    If Len(mLocationError) > 0 Then ws.Range("H1").Value2 = mLocationError
    With ws.UsedRange
        .Font.Name = "Calibri": .Font.Size = 11: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 26
    End With
    ws.Columns("A:E").ColumnWidth = 14: ws.Columns("F").ColumnWidth = 80
    With ws.Range("A1:F1")
        .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True
    End With
End Sub

Private Function ReadStopPatterns(ByVal cfg As Worksheet, ByRef wanted As Object) As Object
    Dim result As Object, data As Variant, i As Long, name As String, tno As String, func As String, voltage As Variant
    Set result = CreateObject("Scripting.Dictionary"): result.CompareMode = vbTextCompare
    Set wanted = CreateObject("Scripting.Dictionary")
    data = cfg.Range("D2:G901").Value2
    For i = LBound(data, 1) To UBound(data, 1)
        If IsError(data(i, 4)) Then Err.Raise vbObjectError + 451, , "Stop Sign cell error at Config G" & i + 1
        name = Trim$(CStr(data(i, 4)))
        If Len(name) > 0 Then
            If IsError(data(i, 1)) Then Err.Raise vbObjectError + 451, , "Function TNO cell error at row " & i + 1
            tno = Trim$(CStr(data(i, 1)))
            If Len(tno) <> 3 Then Err.Raise vbObjectError + 451, , "Stop Sign requires Function TNO in column D; row=" & i + 1
            tno = CStr(CLng(UnsignedValue(tno, 999)))
            If CLng(tno) < 100 Then Err.Raise vbObjectError + 451, , "Stop Function TNO must be 100-999."
            If result.Exists(name) Then Err.Raise vbObjectError + 451, , "Duplicate Stop Sign Pattern Name: " & name
            func = "": voltage = Empty
            If Not IsError(data(i, 2)) Then func = Trim$(CStr(data(i, 2)))
            If Not IsError(data(i, 3)) Then voltage = data(i, 3)
            result.Add name, Array(tno, func, voltage)
            If Not wanted.Exists(tno) Then wanted.Add tno, True
        End If
    Next i
    Set ReadStopPatterns = result
End Function

Private Function LoadMapCounts(ByVal ws As Worksheet, ByVal wanted As Object) As Object
    Dim result As Object, last As Long, start As Long, take As Long, data As Variant, i As Long, b As Long
    Dim wt As String, tno As String, key As String, count As Long
    Set result = CreateObject("Scripting.Dictionary")
    last = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For start = 2 To last Step 512
        take = last - start + 1: If take > 512 Then take = 512
        data = ws.Cells(start, 1).Resize(take, 260).Value2
        For i = 1 To take
            If IsError(data(i, 1)) Or IsError(data(i, 2)) Then Err.Raise vbObjectError + 452, , "Invalid FBC output row " & start + i - 1
            tno = Trim$(CStr(data(i, 2)))
            If wanted.Exists(tno) Then
                wt = UCase$(Trim$(CStr(data(i, 1))))
                If wt <> "WT1" And wt <> "WT2" And wt <> "WT3" Then Err.Raise vbObjectError + 452, , "Invalid WT at FBC output row " & start + i - 1
                key = wt & "|" & tno & "|" & Coordinate(CStr(data(i, 3))) & "|" & Coordinate(CStr(data(i, 4)))
                If result.Exists(key) Then Err.Raise vbObjectError + 452, , "Duplicate FBC WT/TNO/XY: " & key
                count = CountFBCRow(data, i)
                result.Add key, count
            End If
        Next i
    Next start
    Set LoadMapCounts = result
End Function

Private Function CountFBCRow(ByRef data As Variant, ByVal row As Long) As Long
    Dim b As Long, count As Long
    For b = 5 To 260
        If IsError(data(row, b)) Then Err.Raise vbObjectError + 452, , "Invalid FBC value."
        If Len(Trim$(CStr(data(row, b)))) = 0 Or Not IsNumeric(data(row, b)) Then Err.Raise vbObjectError + 452, , "Missing/non-numeric FBC."
        If CDbl(data(row, b)) < 0 Or CDbl(data(row, b)) <> Fix(CDbl(data(row, b))) Then Err.Raise vbObjectError + 452, , "FBC must be non-negative integer."
        If CDbl(data(row, b)) > 0 Then count = count + 1
    Next b
    CountFBCRow = count
End Function

Private Function MapResult(ByVal counts As Object, ByVal errors As Object, ByVal wt As String, ByVal tno As String, ByVal dut As String) As Variant
    Dim key As String
    If errors.Exists(wt) Then MapResult = "Data Error": Exit Function
    key = wt & "|" & tno & "|" & dut
    If counts.Exists(key) Then MapResult = CLng(counts(key)) Else MapResult = "No Data"
End Function

Private Function MapStatus(ByVal result As Variant) As String
    If VarType(result) = vbString Then MapStatus = CStr(result): Exit Function
    If CLng(result) = 0 Then MapStatus = "Pass" Else MapStatus = "Fail"
End Function

Private Function MapColor(ByVal result As Variant) As Long
    Select Case MapStatus(result)
        Case "Pass": MapColor = RGB(220, 252, 231)
        Case "Fail": MapColor = RGB(254, 202, 202)
        Case "No Data": MapColor = RGB(191, 191, 191)
        Case Else: MapColor = RGB(255, 217, 102)
    End Select
End Function

Public Function BLK_BuildFailMaps(ByVal wb As Workbook) As String
    Dim patterns As Object, wanted As Object, counts As Object, coords As Object, errors As Object, used As Object
    Dim locations As Worksheet, summary As Worksheet, map As Worksheet, ws As Worksheet
    Dim data As Variant, item As Variant, name As Variant, detail As Variant, xy As Variant, ordered As Collection
    Dim i As Long, phase As Long, last As Long, minX As Long, maxX As Long, minY As Long, maxY As Long
    Dim width As Long, col As Long, row As Long, origin As Long, sheetName As String, result As Variant
    Dim output() As Variant, n As Long, total As Double, text As String, existing As Boolean
    On Error GoTo Failed
    Set patterns = ReadStopPatterns(wb.Worksheets("Config"), wanted)
    If patterns.Count = 0 Then
        Set used = CreateObject("Scripting.Dictionary")
        PruneOldFailMaps wb, used
        Set summary = wb.Worksheets("BLK Fail Count")
        If summary.AutoFilterMode Then summary.AutoFilterMode = False
        summary.UsedRange.Clear
        summary.Range("A1:I1").Value2 = Array("Pattern", "Stop TNO", "Func.", "Test Voltage", "WT", "X", "Y", "Fail BLK Count", "Status")
        BLK_BuildFailMaps = "Map skipped: enter Pattern Name in Config G beside a Function TNO."
        Exit Function
    End If
    Set locations = wb.Worksheets("Chip Locations")
    If Len(CStr(locations.Range("H1").Value2)) > 0 Then Err.Raise vbObjectError + 453, , CStr(locations.Range("H1").Value2)
    last = locations.Cells(locations.Rows.Count, 1).End(xlUp).Row
    If last < 2 Then Err.Raise vbObjectError + 453, , "Import FBC first to collect Raw Data Chip X/Y locations."
    data = locations.Range("A2:B" & last).Value2
    Set coords = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(data, 1)
        text = Coordinate(CStr(data(i, 1))) & "|" & Coordinate(CStr(data(i, 2)))
        If Not coords.Exists(text) Then coords.Add text, True
    Next i
    Set ordered = SortedDUTs(coords)
    xy = Split(CStr(ordered(1)), "|"): minX = CLng(xy(0)): maxX = minX: minY = CLng(xy(1)): maxY = minY
    For Each item In ordered
        xy = Split(CStr(item), "|")
        If CLng(xy(0)) < minX Then minX = CLng(xy(0))
        If CLng(xy(0)) > maxX Then maxX = CLng(xy(0))
        If CLng(xy(1)) < minY Then minY = CLng(xy(1))
        If CLng(xy(1)) > maxY Then maxY = CLng(xy(1))
    Next item
    If CDbl(maxX) - minX > 99 Or CDbl(maxY) - minY > 99 Then Err.Raise vbObjectError + 453, , "Raw Chip coordinate span exceeds 100; review Chip Locations."
    Set errors = CreateObject("Scripting.Dictionary")
    For phase = 1 To 3
        If CStr(locations.Cells(phase + 1, 5).Value2) <> "OK" Then errors.Add "WT" & phase, True
    Next phase
    Set counts = LoadMapCounts(wb.Worksheets("FBC by BLK"), wanted)
    total = CDbl(patterns.Count) * coords.Count * 3
    Set summary = wb.Worksheets("BLK Fail Count")
    If total + 1 > summary.Rows.Count Then Err.Raise vbObjectError + 453, , "BLK Fail Count exceeds Excel row limit."
    If summary.AutoFilterMode Then summary.AutoFilterMode = False
    summary.UsedRange.Clear
    summary.Range("A1:I1").Value2 = Array("Pattern", "Stop TNO", "Func.", "Test Voltage", "WT", "X", "Y", "Fail BLK Count", "Status")
    Set used = CreateObject("Scripting.Dictionary"): used.CompareMode = vbTextCompare
    width = maxX - minX + 1
    For Each name In patterns
        detail = patterns(name): sheetName = FailMapSheetName(CStr(name), used)
        Set map = Nothing
        On Error Resume Next
        Set map = wb.Worksheets(sheetName)
        On Error GoTo Failed
        If Not map Is Nothing Then
            existing = False
            For i = 1 To map.CustomProperties.Count
                If map.CustomProperties(i).Name = "ASCFailMap" Then existing = True
            Next i
            If Not existing Then Err.Raise vbObjectError + 453, , "Map sheet name collides with an existing non-Map sheet: " & sheetName
            map.UsedRange.UnMerge: map.UsedRange.Clear
        Else
            Set map = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
            map.Name = sheetName: map.CustomProperties.Add Name:="ASCFailMap", Value:="1"
        End If
        With map.Cells(1, 1).Resize(maxY - minY + 6, (width + 2) * 3)
            .Font.Name = "Calibri": .Font.Size = 10: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .ColumnWidth = 11: .RowHeight = 42
        End With
        map.Cells(1, 1).Resize(1, (width + 2) * 3).Merge
        map.Cells(1, 1).Value2 = CStr(name) & " Fail Map | Stop TNO=" & detail(0) & " | " & detail(1) & " | Voltage=" & detail(2)
        With map.Cells(1, 1): .Font.Size = 14: .Font.Bold = True: .Font.Color = vbWhite: .Interior.Color = RGB(23, 50, 77): End With
        map.Cells(2, 1).Resize(1, (width + 2) * 3).Merge
        map.Cells(2, 1).Value2 = "Green: Pass(0) | Red: Fail BLK count | Gray: no selected Stop-TNO data | Yellow: data error | X right, Y down"
        ReDim output(1 To ordered.Count * 3, 1 To 9): n = 0
        For phase = 1 To 3
            origin = 1 + (phase - 1) * (width + 2)
            map.Cells(3, origin + 1).Resize(1, width).Merge
            map.Cells(3, origin + 1).Value2 = "WT" & phase
            map.Cells(3, origin + 1).Font.Bold = True
            For col = minX To maxX: map.Cells(4, origin + 1 + col - minX).Value2 = "X=" & col: Next col
            For row = minY To maxY: map.Cells(5 + row - minY, origin).Value2 = "Y=" & row: Next row
        Next phase
        For Each item In ordered
            xy = Split(CStr(item), "|")
            For phase = 1 To 3
                result = MapResult(counts, errors, "WT" & phase, CStr(detail(0)), CStr(item))
                n = n + 1: output(n, 1) = name: output(n, 2) = CLng(detail(0)): output(n, 3) = detail(1): output(n, 4) = detail(2)
                output(n, 5) = "WT" & phase: output(n, 6) = CLng(xy(0)): output(n, 7) = CLng(xy(1)): output(n, 9) = MapStatus(result)
                If VarType(result) <> vbString Then output(n, 8) = result
                origin = 1 + (phase - 1) * (width + 2)
                With map.Cells(5 + CLng(xy(1)) - minY, origin + 1 + CLng(xy(0)) - minX)
                    .Value2 = "(" & xy(0) & "," & xy(1) & ")" & vbLf & CStr(result)
                    .WrapText = True: .Interior.Color = MapColor(result)
                    .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(150, 150, 150)
                End With
            Next phase
        Next item
        last = summary.Cells(summary.Rows.Count, 1).End(xlUp).Row + 1
        summary.Cells(last, 1).Resize(n, 9).Value2 = output
        For i = 1 To n
            If output(i, 9) = "Pass" Then
                result = 0
            ElseIf output(i, 9) = "Fail" Then
                result = output(i, 8)
            Else
                result = output(i, 9)
            End If
            summary.Cells(last + i - 1, 1).Resize(1, 9).Interior.Color = MapColor(result)
        Next i
    Next name
    With summary.Cells(1, 1).Resize(CLng(total) + 1, 9)
        .Font.Name = "Calibri": .Font.Size = 11: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 26
    End With
    summary.Columns("A:D").ColumnWidth = 20: summary.Columns("E:G").ColumnWidth = 10: summary.Columns("H:I").ColumnWidth = 18
    summary.Columns("D").NumberFormat = "0.000"
    With summary.Range("A1:I1")
        .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: .RowHeight = 32
    End With
    summary.Cells(1, 1).Resize(CLng(total) + 1, 9).AutoFilter
    PruneOldFailMaps wb, used
    Exit Function
Failed:
    BLK_BuildFailMaps = "Map build failed (" & Err.Number & "): " & Err.Description
End Function

Private Sub PruneOldFailMaps(ByVal wb As Workbook, ByVal keep As Object)
    Dim ws As Worksheet, i As Long, j As Long, managed As Boolean, oldAlerts As Boolean, description As String
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    Application.DisplayAlerts = False
    For i = wb.Worksheets.Count To 1 Step -1
        Set ws = wb.Worksheets(i): managed = False
        For j = 1 To ws.CustomProperties.Count
            If ws.CustomProperties(j).Name = "ASCFailMap" Then managed = True
        Next j
        If managed And Not keep.Exists(ws.Name) Then ws.Delete
    Next i
    Application.DisplayAlerts = oldAlerts
    Exit Sub
Failed:
    description = Err.Description
    Application.DisplayAlerts = oldAlerts
    Err.Raise vbObjectError + 454, , "Old Map cleanup: " & description
End Sub

Private Function FailMapSheetName(ByVal pattern As String, ByVal used As Object) As String
    Dim clean As String, bad As Variant, name As String, i As Long, suffix As String
    clean = Trim$(pattern)
    For Each bad In Array("[", "]", ":", "*", "?", "/", "\", "'"): clean = Replace(clean, CStr(bad), "_"): Next bad
    For i = 0 To 31: clean = Replace(clean, Chr$(i), "_"): Next i
    name = Left$(clean, 22) & " Fail Map": i = 1
    Do While used.Exists(name)
        i = i + 1: suffix = "~" & i
        name = Left$(clean, 22 - Len(suffix)) & suffix & " Fail Map"
    Loop
    used.Add name, True: FailMapSheetName = name
End Function
