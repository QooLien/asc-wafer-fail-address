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
    mCarry = "": mActive = False
End Sub

Public Sub BLK_BeginFile(ByVal path As String)
    Dim i As Long
    Set mWanted = CreateObject("Scripting.Dictionary")
    For i = LBound(mSequence) To UBound(mSequence): mWanted.Add mSequence(i), True: Next i
    ResetParser path
    mCarry = "": mBLKError = "": mIgnored = 0
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


Private Sub ResetParser(ByVal path As String)
    Set mData = CreateObject("Scripting.Dictionary"): Set mDUTs = CreateObject("Scripting.Dictionary")
    mTno = "": mX = "": mY = "": mPath = path: mLine = 0
    mActive = False: mCount = 0
    If mTags Is Nothing Then
        Set mTags = CreateObject("VBScript.RegExp"): mTags.Global = True: mTags.IgnoreCase = True
        mTags.Pattern = "\b(Tno|X|Y)\s*=\s*([+-]?[0-9]+)(?=\s|[,;]|$)"
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
