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
Private mAvailablePhases As Long
Private mCats As Object, mCatTag As Object
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
    mNextRow = 2: mAvailablePhases = 0
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
    BLK_SaveCategories wb, phase
    If phase > mAvailablePhases Then mAvailablePhases = phase
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
    ws.Cells(startRow, 5).Resize(mNextRow - startRow, 256).NumberFormat = "0;-0;;@"
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
    BLK_Finalize = BLK_BuildComparison(wb)
    Exit Function
Failed:
    BLK_Finalize = "FBC final formatting: " & Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Sort.SortFields.Clear
        If lastRow > 0 Then ws.Cells(1, 261).Resize(lastRow, 1).ClearContents
    End If
End Function

Public Function BLK_BuildComparison(ByVal wb As Workbook) As String
    Dim cfg As Worksheet, source As Worksheet, locations As Worksheet, ws As Worksheet
    Dim tnos As Object, labels As Object, coords As Object, spans As Object, seen As Object, types As Object
    Dim settings As Variant, names As Variant, data As Variant, span As Variant, xy As Variant, record As Variant
    Dim ordered As Collection, item As Variant, tno As String, key As String, dut As String, wt As String
    Dim cols As Variant, col As Variant, i As Long, r As Long, b As Long, phase As Long, n As Long
    Dim last As Long, start As Long, take As Long, width As Long, target As Long, dest As Long, color As Long
    Dim header() As Variant, packet() As Variant, info() As Variant, infoRow As Long, mapping As Worksheet
    Dim phaseCount As Long, monitorCount As Long, group As Long, groupCount As Long, firstIndex As Long, groupBase As Long, groupName As String, columnMap As Object
    On Error GoTo Failed
    Set cfg = wb.Worksheets("Config"): Set source = wb.Worksheets("FBC by BLK")
    Set locations = wb.Worksheets("Chip Locations")
    Set tnos = CreateObject("Scripting.Dictionary"): Set types = CreateObject("Scripting.Dictionary")
    cols = Array("A2:A901", "D2:D901")
    For Each col In cols
        settings = cfg.Range(CStr(col)).Value2
        For i = 1 To UBound(settings, 1)
            If IsError(settings(i, 1)) Then Err.Raise vbObjectError + 490, , "Invalid Config TNO."
            tno = Trim$(CStr(settings(i, 1)))
            If Len(tno) > 0 Then
                If Len(tno) <> 3 Then Err.Raise vbObjectError + 490, , "Config TNO must be100-999."
                tno = CStr(CLng(UnsignedValue(tno, 999)))
                If CLng(tno) < 100 Or tnos.Exists(tno) Then Err.Raise vbObjectError + 490, , "Invalid/duplicate comparison TNO: " & tno
                tnos.Add tno, tnos.Count
                If CStr(col) = "A2:A901" Then
                    types.Add tno, "Monitor": monitorCount = monitorCount + 1
                Else
                    types.Add tno, "Function"
                End If
            End If
        Next i
    Next col
    If tnos.Count = 0 Then Err.Raise vbObjectError + 490, , "Enter Config TNOs before comparison."
    Set labels = ImportFailAddresses.BuildAddressLabels(cfg)
    phaseCount = BLK_PhaseCount(wb)
    width = 3 + phaseCount * tnos.Count
    Set coords = CreateObject("Scripting.Dictionary"): Set spans = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary")
    last = PresentationLastRow(locations, 1)
    If last >= 2 Then
        data = locations.Range("A2:B" & last).Value2
        For i = 1 To UBound(data, 1)
            coords(Coordinate(CStr(data(i, 1))) & "|" & Coordinate(CStr(data(i, 2)))) = True
        Next i
    End If
    Set ordered = SortedDUTs(coords)
    If CDbl(ordered.Count) * 256 + 2 > source.Rows.Count Then Err.Raise vbObjectError + 490, , "BLK comparison exceeds Excel row limit."
    last = PresentationLastRow(source, 2)
    For start = 2 To last Step 512
        take = last - start + 1: If take > 512 Then take = 512
        data = source.Cells(start, 1).Resize(take, 4).Value2
        For r = 1 To take
            tno = CStr(data(r, 2))
            If tnos.Exists(tno) Then
                wt = UCase$(Trim$(CStr(data(r, 1))))
                If wt <> "WT1" And wt <> "WT2" And wt <> "WT3" Then Err.Raise vbObjectError + 490, , "Invalid comparison WT."
                dut = Coordinate(CStr(data(r, 3))) & "|" & Coordinate(CStr(data(r, 4)))
                key = wt & "|" & tno & "|" & dut
                If seen.Exists(key) Then Err.Raise vbObjectError + 490, , "Duplicate comparison WT/TNO/Chip."
                seen.Add key, True
                If spans.Exists(dut) Then
                    span = spans(dut): span(1) = start + r - 1: spans(dut) = span
                Else
                    spans.Add dut, Array(start + r - 1, start + r - 1)
                End If
            End If
        Next r
    Next start
    On Error Resume Next
    Set ws = wb.Worksheets("BLK_FBC_Comparison")
    If ws Is Nothing Then Set ws = wb.Worksheets(ChrW(24037) & ChrW(20316) & ChrW(34920) & "1")
    On Error GoTo Failed
    If ws Is Nothing Then Set ws = wb.Worksheets.Add(After:=source)
    ws.Name = "BLK_FBC_Comparison": ws.Visible = xlSheetVisible
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.UsedRange.UnMerge: ws.UsedRange.Clear: ws.UsedRange.ClearComments
    If ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1 > width Then ws.Range(ws.Columns(width + 1), ws.Columns(ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1)).Delete
    If ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1 > ordered.Count * 256 + 2 Then ws.Range(ws.Rows(ordered.Count * 256 + 3), ws.Rows(ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1)).Delete
    ReDim header(1 To 1, 1 To width)
    header(1, 1) = "X": header(1, 2) = "Y": header(1, 3) = "BLK"
    names = tnos.Keys
    ReDim info(1 To phaseCount * tnos.Count, 1 To 5)
    With ws.Cells(1, 1).Resize(2 + 256 * ordered.Count, width)
        .Font.Name = "Calibri": .Font.Size = 11: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .RowHeight = 26: .ColumnWidth = 17
    End With
    ws.Columns("A:C").ColumnWidth = 10
    ws.Range("A1:C1").Merge: ws.Range("A1").Value2 = "Chip / BLK"
    Set columnMap = CreateObject("Scripting.Dictionary")
    For group = 0 To 1
        If group = 0 Then
            groupName = "Monitor": groupCount = monitorCount: firstIndex = 0: groupBase = 4
        Else
            groupName = "Function": groupCount = tnos.Count - monitorCount: firstIndex = monitorCount: groupBase = 4 + phaseCount * monitorCount
        End If
        If groupCount > 0 Then
            For phase = 1 To phaseCount
                target = groupBase + (phase - 1) * groupCount
                Select Case phase
                    Case 1: color = RGB(219, 234, 254)
                    Case 2: color = RGB(255, 237, 213)
                    Case 3: color = RGB(220, 252, 231)
                End Select
                ws.Cells(1, target).Resize(1, groupCount).Merge
                ws.Cells(1, target).Value2 = "WT" & phase & " / " & groupName
                ws.Cells(1, target).Resize(2, groupCount).Interior.Color = color
                For i = 0 To groupCount - 1
                    tno = CStr(names(firstIndex + i)): header(1, target + i) = labels(tno)
                    columnMap.Add "WT" & phase & "|" & tno, target + i
                    ws.Cells(2, target + i).Orientation = 90
                    ws.Cells(2, target + i).AddComment "WT" & phase & " | " & groupName & " | TNO=" & tno
                    infoRow = infoRow + 1: info(infoRow, 1) = "WT" & phase: info(infoRow, 2) = CLng(tno)
                    info(infoRow, 3) = target + i: info(infoRow, 4) = groupName: info(infoRow, 5) = labels(tno)
                Next i
            Next phase
        End If
    Next group
    ws.Cells(2, 1).Resize(1, width).Value2 = header
    With ws.Cells(1, 1).Resize(2, width)
        .Font.Bold = True: .WrapText = True: .RowHeight = 42: .Borders.LineStyle = xlContinuous
    End With
    ws.Range("A1:C2").Interior.Color = RGB(23, 50, 77): ws.Range("A1:C2").Font.Color = vbWhite
    dest = 3
    For Each item In ordered
        dut = CStr(item): xy = Split(dut, "|")
        ReDim packet(1 To 256, 1 To width)
        For b = 0 To 255
            packet(b + 1, 1) = CLng(xy(0)): packet(b + 1, 2) = CLng(xy(1)): packet(b + 1, 3) = b
        Next b
        ws.Cells(dest, 4).Resize(256, width - 3).Interior.Color = RGB(224, 224, 224)
        If spans.Exists(dut) Then
            span = spans(dut)
            For start = CLng(span(0)) To CLng(span(1)) Step 512
                take = CLng(span(1)) - start + 1: If take > 512 Then take = 512
                data = source.Cells(start, 1).Resize(take, 260).Value2
                For r = 1 To take
                    key = Coordinate(CStr(data(r, 3))) & "|" & Coordinate(CStr(data(r, 4)))
                    tno = CStr(data(r, 2))
                    If key = dut And tnos.Exists(tno) Then
                        phase = CLng(Right$(CStr(data(r, 1)), 1))
                        If CStr(locations.Cells(phase + 1, 5).Value2) = "OK" Then
                            n = CountFBCRow(data, r)
                            target = CLng(columnMap("WT" & phase & "|" & tno))
                            For b = 0 To 255: packet(b + 1, target) = CDbl(data(r, b + 5)): Next b
                            ws.Cells(dest, target).Resize(256, 1).Interior.Color = vbWhite
                        End If
                    End If
                Next r
            Next start
        End If
        ws.Cells(dest, 1).Resize(256, width).Value2 = packet
        ws.Cells(dest, 3).Resize(256, 1).NumberFormat = """BLK""0"
        ws.Cells(dest, 4).Resize(256, width - 3).NumberFormat = "0;-0;;@"
        ws.Cells(dest, 1).Resize(256, width).BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(80, 80, 80)
        dest = dest + 256
        Application.StatusBar = "BLK FBC comparison: " & dut
    Next item
    ws.Cells(2, 1).Resize(dest - 2, width).AutoFilter
    ws.Cells(1, 1).Resize(dest - 1, width).BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(80, 80, 80)
    Set mapping = wb.Worksheets("Import Index")
    mapping.Range("Q:U").ClearContents
    mapping.Range("Q1:U1").Value2 = Array("Comparison WT", "TNO", "Column", "Group", "Header")
    mapping.Range("Q2").Resize(infoRow, 5).Value2 = info
    With mapping.UsedRange
        .Font.Name = "Calibri": .Font.Size = 11: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    BLK_SizeOutput ws
    ws.Activate
    With Application.ActiveWindow
        .DisplayGridlines = False: .FreezePanes = False: .SplitColumn = 3: .SplitRow = 2: .FreezePanes = True
    End With
    source.Visible = xlSheetVeryHidden
    Exit Function
Failed:
    BLK_BuildComparison = "FBC comparison failed (" & Err.Number & "): " & Err.Description
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
    mCarry = "": mActive = False: mAvailablePhases = 0
End Sub

Public Sub BLK_BeginFile(ByVal path As String)
    Dim i As Long
    Set mWanted = CreateObject("Scripting.Dictionary")
    For i = LBound(mSequence) To UBound(mSequence): mWanted.Add mSequence(i), True: Next i
    Set mCats = CreateObject("Scripting.Dictionary")
    ResetParser path
    mCarry = "": mBLKError = "": mIgnored = 0: mLocationCarry = ""
End Sub

Public Sub BLK_FeedCategories(ByVal raw As String, ByVal atEOF As Boolean)
    Dim pos As Long, line As String
    mCarry = mCarry & raw
    Do
        pos = InStr(mCarry, vbLf)
        If pos = 0 Then Exit Do
        line = Left$(mCarry, pos - 1): mCarry = Mid$(mCarry, pos + 1)
        mLine = mLine + 1
        If InStr(1, line, "CATLST", vbTextCompare) > 0 Then ParseCategory NormalizeASCWhitespace(Replace(line, vbCr, ""))
    Loop
    If Len(mCarry) > MAX_LINE Then Err.Raise vbObjectError + 310, , "ASC line exceeds 1 MB."
    If atEOF Then
        If Len(mCarry) > 0 Then
            mLine = mLine + 1
            If InStr(1, mCarry, "CATLST", vbTextCompare) > 0 Then ParseCategory NormalizeASCWhitespace(Replace(mCarry, vbCr, ""))
        End If
        mCarry = ""
    End If
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
    ParseCategory line
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

Private Sub ParseCategory(ByVal line As String)
    Dim hits As Object, tags As Object, bits As String, pos As Long, key As String, cat As Long
    If InStr(1, line, "CATLST", vbTextCompare) = 0 Then Exit Sub
    If mCatTag Is Nothing Then
        Set mCatTag = CreateObject("VBScript.RegExp")
        mCatTag.IgnoreCase = True: mCatTag.Pattern = "\bCATLST\s*:\s*([01]+)(?=\s|$)"
    End If
    Set hits = mCatTag.Execute(line)
    If hits.Count <> 1 Then Err.Raise vbObjectError + 319, , "Invalid CATLST; line=" & mLine
    Set tags = mTags.Execute(Left$(line, hits(0).FirstIndex))
    If RecordHeaderInteger(tags, "TNO") <> "1000" Then Exit Sub
    key = RecordHeaderInteger(tags, "X") & "|" & RecordHeaderInteger(tags, "Y")
    bits = hits(0).SubMatches(0): cat = InStr(1, bits, "1", vbBinaryCompare)
    If cat = 0 Then Exit Sub
    If mCats.Exists(key) Then
        If CLng(mCats(key)) <> cat Then Err.Raise vbObjectError + 319, , "Conflicting CATLST for Chip " & key
    Else
        mCats.Add key, cat
    End If
End Sub

Public Sub BLK_SaveCategories(ByVal wb As Workbook, ByVal phase As Long)
    Dim ws As Worksheet, data() As Variant, key As Variant, xy As Variant, i As Long, start As Long
    Set ws = wb.Worksheets("Import Index")
    If phase = 1 Then ws.Range("AA:AD").ClearContents
    ws.Range("AA1:AD1").Value2 = Array("WT", "X", "Y", "Cat.")
    If mCats.Count = 0 Then Exit Sub
    ReDim data(1 To mCats.Count, 1 To 4)
    For Each key In mCats
        i = i + 1: xy = Split(CStr(key), "|")
        data(i, 1) = "WT" & phase: data(i, 2) = CLng(xy(0)): data(i, 3) = CLng(xy(1)): data(i, 4) = mCats(key)
    Next key
    start = ws.Cells(ws.Rows.Count, 27).End(xlUp).Row + 1
    ws.Cells(start, 27).Resize(i, 4).Value2 = data
End Sub

Public Function BLK_Categories(ByVal wb As Workbook, Optional ByVal wt As String = "WT1") As Object
    Dim result As Object, ws As Worksheet, last As Long, data As Variant, i As Long, key As String
    Set result = CreateObject("Scripting.Dictionary")
    Set ws = wb.Worksheets("Import Index"): last = ws.Cells(ws.Rows.Count, 27).End(xlUp).Row
    If last >= 2 Then
        data = ws.Range("AA2:AD" & last).Value2
        For i = 1 To UBound(data, 1)
            If wt = "" Or CStr(data(i, 1)) = wt Then
                key = CStr(data(i, 2)) & "|" & CStr(data(i, 3))
                result(key) = CLng(data(i, 4))
            End If
        Next i
    End If
    Set BLK_Categories = result
End Function

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
    For i = 1 To mAvailablePhases
        ws.Cells(i + 1, 4).Value2 = "WT" & i
        If mPhaseErrors.Exists("WT" & i) Then
            If Len(CStr(mPhaseErrors("WT" & i))) > 0 Then ws.Cells(i + 1, 5).Value2 = "Data Error" Else ws.Cells(i + 1, 5).Value2 = "OK"
            ws.Cells(i + 1, 6).Value2 = mPhaseErrors("WT" & i)
        Else
            ws.Cells(i + 1, 5).Value2 = "Not tested": ws.Cells(i + 1, 6).Value2 = ""
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

Private Function ReadStopPatterns(ByVal cfg As Worksheet, ByRef wanted As Object, ByRef hard As Object) As Object
    Dim result As Object, data As Variant, i As Long, name As String, tno As String, func As String, voltage As Variant
    Set result = CreateObject("Scripting.Dictionary"): result.CompareMode = vbTextCompare
    Set wanted = CreateObject("Scripting.Dictionary")
    Set hard = CreateObject("Scripting.Dictionary")
    data = cfg.Range("D2:G901").Value2
    For i = LBound(data, 1) To UBound(data, 1)
        If IsError(data(i, 1)) Then Err.Raise vbObjectError + 451, , "Function TNO cell error at row " & i + 1
        tno = Trim$(CStr(data(i, 1)))
        If Len(tno) > 0 Then
            If Len(tno) <> 3 Then Err.Raise vbObjectError + 451, , "Function TNO must be 100-999; row=" & i + 1
            tno = CStr(CLng(UnsignedValue(tno, 999)))
            If CLng(tno) < 100 Then Err.Raise vbObjectError + 451, , "Function TNO must be 100-999."
            If Not wanted.Exists(tno) Then wanted.Add tno, True
        End If
        If IsError(data(i, 4)) Then Err.Raise vbObjectError + 451, , "Stop Sign cell error at Config G" & i + 1
        name = Trim$(CStr(data(i, 4)))
        If Len(name) > 0 Then
            If IsError(data(i, 1)) Then Err.Raise vbObjectError + 451, , "Function TNO cell error at row " & i + 1
            tno = Trim$(CStr(data(i, 1)))
            If Len(tno) <> 3 Then Err.Raise vbObjectError + 451, , "Stop Sign requires Function TNO in column D; row=" & i + 1
            tno = CStr(CLng(UnsignedValue(tno, 999)))
            If CLng(tno) < 100 Then Err.Raise vbObjectError + 451, , "Stop Function TNO must be 100-999."
            If StrComp(name, "Hard Fail", vbTextCompare) <> 0 Then
            If result.Exists(name) Then Err.Raise vbObjectError + 451, , "Duplicate Stop Sign Pattern Name: " & name
            func = "": voltage = Empty
            If Not IsError(data(i, 2)) Then func = Trim$(CStr(data(i, 2)))
            If Not IsError(data(i, 3)) Then voltage = data(i, 3)
            result.Add name, Array(tno, func, voltage)
            Else
                If Not hard.Exists(tno) Then hard.Add tno, True
            End If
            If Not wanted.Exists(tno) Then wanted.Add tno, True
        End If
    Next i
    Set ReadStopPatterns = result
End Function

Private Function PatternPredecessors(ByVal patterns As Object) As Object
    'ReadStopPatterns preserves Config row order. Numeric TNO order is irrelevant.
    Dim result As Object, earlier As New Collection, required As Collection
    Dim name As Variant, detail As Variant, tno As Variant
    Set result = CreateObject("Scripting.Dictionary")
    For Each name In patterns
        detail = patterns(name): Set required = New Collection
        For Each tno In earlier
            If CStr(tno) <> CStr(detail(0)) Then required.Add CStr(tno)
        Next tno
        If Not result.Exists(CStr(detail(0))) Then result.Add CStr(detail(0)), required
        earlier.Add CStr(detail(0))
    Next name
    Set PatternPredecessors = result
End Function

Private Function LoadMapCounts(ByVal ws As Worksheet, ByVal wanted As Object) As Object
    Dim result As Object, last As Long, start As Long, take As Long, data As Variant, i As Long, b As Long
    Dim wt As String, tno As String, key As String, count As Long
    Set result = CreateObject("Scripting.Dictionary")
    last = PresentationLastRow(ws, 2)
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
                result.Add key, FBCRowMask(data, i)
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

Private Function FBCRowMask(ByRef data As Variant, ByVal row As Long) As String
    Dim b As Long, mask As String
    mask = String$(256, "0")
    For b = 0 To 255
        If CDbl(data(row, b + 5)) > 0 Then Mid$(mask, b + 1, 1) = "1"
    Next b
    FBCRowMask = mask
End Function

Private Function MapResult(ByVal counts As Object, ByVal errors As Object, ByVal wt As String, ByVal tno As String, ByVal dut As String, ByVal hard As Object, ByVal functionTnos As Object, Optional ByRef includedMask As Variant, Optional ByVal priorPatterns As Object = Nothing) As Variant
    Dim key As String, target As String, blocked As String, mask As String
    Dim phase As Long, currentPhase As Long, b As Long, n As Long, h As Variant, f As Variant
    Dim required As Collection
    includedMask = ""
    currentPhase = CLng(Right$(wt, 1))
    key = wt & "|" & tno & "|" & dut
    If errors.Exists(wt) Then MapResult = "Data Error": Exit Function
    If Not counts.Exists(key) Then MapResult = "No Data": Exit Function
    target = CStr(counts(key)): blocked = String$(256, "0")
    'Hard Fail is permanent for this BLK: include current and all preceding WTs.
    For phase = 1 To currentPhase
        If errors.Exists("WT" & phase) Then MapResult = "Data Error": Exit Function
        For Each h In hard
            key = "WT" & phase & "|" & CStr(h) & "|" & dut
            If Not counts.Exists(key) Then MapResult = "No Data": Exit Function
            mask = CStr(counts(key))
            For b = 1 To 256
                If Mid$(mask, b, 1) = "1" Then Mid$(blocked, b, 1) = "1"
            Next b
        Next h
        'Every configured Function TNO in every earlier WT must Pass at this Chip/BLK.
        If phase < currentPhase Then
            For Each f In functionTnos
            key = "WT" & phase & "|" & CStr(f) & "|" & dut
            If Not counts.Exists(key) Then MapResult = "No Data": Exit Function
            mask = CStr(counts(key))
            For b = 1 To 256
                If Mid$(mask, b, 1) = "1" Then Mid$(blocked, b, 1) = "1"
            Next b
            Next f
        End If
    Next phase
    'Within this WT, every earlier Pattern endpoint must Pass at this Chip/BLK.
    'No Address Count is read here: the Map depends only on FBC presence/values.
    If Not priorPatterns Is Nothing Then
        If priorPatterns.Exists(tno) Then
            Set required = priorPatterns(tno)
            For Each f In required
                key = wt & "|" & CStr(f) & "|" & dut
                If Not counts.Exists(key) Then MapResult = "No Data": Exit Function
                mask = CStr(counts(key))
                For b = 1 To 256
                    If Mid$(mask, b, 1) = "1" Then Mid$(blocked, b, 1) = "1"
                Next b
            Next f
        End If
    End If
    includedMask = String$(256, "0")
    For b = 1 To 256
        If Mid$(blocked, b, 1) = "0" And Mid$(target, b, 1) = "1" Then
            n = n + 1: Mid$(includedMask, b, 1) = "1"
        End If
    Next b
    MapResult = n
End Function

Public Function BLK_MapIncludedMasks(ByVal wb As Workbook) As Object
    'Audit uses exactly the same eligibility evaluator as the Map, without changing sheets.
    Dim patterns As Object, wanted As Object, hard As Object, counts As Object, errors As Object, endpoints As Object, result As Object
    Dim name As Variant, detail As Variant, key As Variant, parts As Variant, phase As Long, value As Variant, mask As Variant
    Dim locations As Worksheet, priorPatterns As Object
    Dim coords As Object, locationData As Variant, last As Long, i As Long
    Set patterns = ReadStopPatterns(wb.Worksheets("Config"), wanted, hard)
    If patterns.Count = 0 Then Err.Raise vbObjectError + 486, , "No Pattern endpoints in Config G. Hard Fail alone does not define a Map."
    Set priorPatterns = PatternPredecessors(patterns)
    Set endpoints = CreateObject("Scripting.Dictionary")
    For Each name In patterns
        detail = patterns(name)
        If Not endpoints.Exists(CStr(detail(0))) Then endpoints.Add CStr(detail(0)), True
    Next name
    Set errors = CreateObject("Scripting.Dictionary")
    Set locations = wb.Worksheets("Chip Locations")
    If Len(CStr(locations.Range("H1").Value2)) > 0 Then Err.Raise vbObjectError + 486, , CStr(locations.Range("H1").Value2)
    Set coords = CreateObject("Scripting.Dictionary")
    last = locations.Cells(locations.Rows.Count, 1).End(xlUp).Row
    If last < 2 Then Err.Raise vbObjectError + 486, , "Import FBC first to collect Raw Chip locations."
    locationData = locations.Range("A2:B" & last).Value2
    For i = 1 To UBound(locationData, 1)
        coords(Coordinate(CStr(locationData(i, 1))) & "|" & Coordinate(CStr(locationData(i, 2)))) = True
    Next i
    For phase = 1 To 3
        If CStr(locations.Cells(phase + 1, 5).Value2) = "Data Error" Then errors.Add "WT" & phase, True
    Next phase
    Set counts = LoadMapCounts(wb.Worksheets("FBC by BLK"), wanted)
    Set result = CreateObject("Scripting.Dictionary")
    For Each key In counts
        parts = Split(CStr(key), "|")
        If endpoints.Exists(CStr(parts(1))) And coords.Exists(CStr(parts(2)) & "|" & CStr(parts(3))) Then
            value = MapResult(counts, errors, CStr(parts(0)), CStr(parts(1)), CStr(parts(2)) & "|" & CStr(parts(3)), hard, wanted, mask, priorPatterns)
            If VarType(value) <> vbString Then result.Add CStr(key), CStr(mask)
        End If
    Next key
    Set BLK_MapIncludedMasks = result
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

Private Sub SetMapDrillProperty(ByVal ws As Worksheet, ByVal key As String, ByVal value As String)
    Dim i As Long
    For i = ws.CustomProperties.Count To 1 Step -1
        If ws.CustomProperties(i).Name = key Then ws.CustomProperties(i).Delete
    Next i
    ws.CustomProperties.Add Name:=key, Value:=value
End Sub

Private Sub ApplyMapColorScale(ByVal ws As Worksheet, ByVal minX As Long, ByVal maxX As Long, ByVal minY As Long, ByVal maxY As Long, ByVal phaseCount As Long, ByVal maximum As Long)
    Dim phase As Long, origin As Long, width As Long, cells As Range
    Dim colorScaleRule As Object, zeroRule As Object, first As String
    width = maxX - minX + 1
    If maximum < 2 Then maximum = 2
    For phase = 1 To phaseCount
        origin = 1 + (phase - 1) * (width + 2)
        Set cells = ws.Cells(5, origin + 1).Resize(maxY - minY + 1, width)
        cells.FormatConditions.Delete
        Set colorScaleRule = cells.FormatConditions.AddColorScale(ColorScaleType:=2)
        With colorScaleRule.ColorScaleCriteria(1)
            .Type = xlConditionValueNumber: .Value = 1: .FormatColor.Color = RGB(255, 235, 132)
        End With
        With colorScaleRule.ColorScaleCriteria(2)
            .Type = xlConditionValueNumber: .Value = maximum: .FormatColor.Color = RGB(248, 105, 107)
        End With
        first = cells.Cells(1, 1).Address(False, False)
        Set zeroRule = cells.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISNUMBER(" & first & ")," & first & "=0)")
        zeroRule.SetFirstPriority: zeroRule.StopIfTrue = True
        zeroRule.Interior.Color = vbWhite: zeroRule.Font.Color = vbWhite
    Next phase
End Sub


Private Function ChipFailEvidence(ByVal wb As Workbook) As Object
    Dim result As Object, ws As Worksheet, last As Long, start As Long, take As Long
    Dim data As Variant, i As Long, b As Long, key As String
    Set result = ImportFailAddresses.MapChipAddressEvidence(wb)
    Set ws = wb.Worksheets("FBC by BLK"): last = PresentationLastRow(ws, 2)
    For start = 2 To last Step 512
        take = last - start + 1: If take > 512 Then take = 512
        data = ws.Cells(start, 1).Resize(take, 260).Value2
        For i = 1 To take
            key = CStr(data(i, 1)) & "|" & CStr(data(i, 3)) & "|" & CStr(data(i, 4))
            If Not result.Exists(key) Then
                For b = 5 To 260
                    If Not IsError(data(i, b)) Then
                        If IsNumeric(data(i, b)) Then
                            If CDbl(data(i, b)) > 0 Then result.Add key, True: Exit For
                        End If
                    End If
                Next b
            End If
        Next i
    Next start
    Set ChipFailEvidence = result
End Function

Public Function BLK_BuildFailMaps(ByVal wb As Workbook) As String
    Dim patterns As Object, wanted As Object, counts As Object, coords As Object, errors As Object, used As Object
    Dim phaseCount As Long, includedMask As Variant, maximumCount As Long
    Dim priorPatterns As Object, categoryByWT(1 To 3) As Object, chipCat As Long, evidence As Object
    Dim hard As Object
    Dim locations As Worksheet, summary As Worksheet, map As Worksheet, ws As Worksheet
    Dim data As Variant, item As Variant, name As Variant, detail As Variant, xy As Variant, ordered As Collection
    Dim i As Long, phase As Long, last As Long, minX As Long, maxX As Long, minY As Long, maxY As Long
    Dim width As Long, col As Long, row As Long, origin As Long, sheetName As String, result As Variant
    Dim output() As Variant, n As Long, total As Double, text As String, existing As Boolean
    On Error GoTo Failed
    RevealMapSources wb, True
    BLK_InvalidateCoverageMap wb
    BLK_ClearChipMonitorMap wb
    BLK_InvalidateDetails wb, "Map inputs updated. Double-click a Fail Chip to refresh this detail view."
    phaseCount = BLK_PhaseCount(wb)
    Set patterns = ReadStopPatterns(wb.Worksheets("Config"), wanted, hard)
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
    For phase = 1 To phaseCount
        If CStr(locations.Cells(phase + 1, 5).Value2) = "Data Error" Then errors.Add "WT" & phase, True
    Next phase
    For phase = 1 To phaseCount: Set categoryByWT(phase) = BLK_Categories(wb, "WT1"): Next phase
    Set priorPatterns = PatternPredecessors(patterns)
    Set counts = LoadMapCounts(wb.Worksheets("FBC by BLK"), wanted)
    Set evidence = ChipFailEvidence(wb)
    total = CDbl(patterns.Count) * coords.Count * phaseCount
    Set summary = wb.Worksheets("BLK Fail Count")
    If total + 1 > summary.Rows.Count Then Err.Raise vbObjectError + 453, , "BLK Fail Count exceeds Excel row limit."
    If summary.AutoFilterMode Then summary.AutoFilterMode = False
    summary.UsedRange.Clear
    summary.Range("A1:I1").Value2 = Array("Pattern", "Stop TNO", "Func.", "Test Voltage", "WT", "X", "Y", "Fail BLK Count", "Status")
    Set used = CreateObject("Scripting.Dictionary"): used.CompareMode = vbTextCompare
    width = maxX - minX + 1
    For Each name In patterns
        detail = patterns(name): sheetName = FailMapSheetName(CStr(name), used): maximumCount = 0
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
            If map.UsedRange.Column + map.UsedRange.Columns.Count - 1 > (width + 2) * phaseCount Then map.Range(map.Columns((width + 2) * phaseCount + 1), map.Columns(map.UsedRange.Column + map.UsedRange.Columns.Count - 1)).Delete
        Else
            Set map = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
            map.Name = sheetName: map.CustomProperties.Add Name:="ASCFailMap", Value:="1"
        End If
        SetMapDrillProperty map, "ASCPattern", CStr(name)
        SetMapDrillProperty map, "ASCStopTNO", CStr(detail(0))
        With map.Cells(1, 1).Resize(maxY - minY + 6, (width + 2) * phaseCount)
            .Font.Name = "Calibri": .Font.Size = 24: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
            .ColumnWidth = 12: .RowHeight = 40
        End With
        map.Cells(1, 1).Resize(1, (width + 2) * phaseCount).Merge
        map.Cells(1, 1).Value2 = CStr(name) & " Fail Map"
        With map.Cells(1, 1): .Font.Size = 24: .Font.Bold = True: .Font.Color = vbWhite: .Interior.Color = RGB(23, 50, 77): End With
        map.Cells(2, 1).Resize(1, (width + 2) * phaseCount).Merge
        map.Rows(2).RowHeight = 12
        ReDim output(1 To ordered.Count * phaseCount, 1 To 9): n = 0
        For phase = 1 To phaseCount
            origin = 1 + (phase - 1) * (width + 2)
            map.Cells(3, origin + 1).Resize(1, width).Merge
            map.Cells(3, origin + 1).Value2 = "WT" & phase
            map.Cells(3, origin + 1).Font.Bold = True
            map.Cells(3, origin + 1).Resize(1, width).BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(80, 80, 80)
            map.Cells(4, origin + 1).Resize(1, width).Borders.LineStyle = xlContinuous
            map.Cells(5, origin).Resize(maxY - minY + 1, 1).Borders.LineStyle = xlContinuous
            map.Cells(3, origin).Resize(maxY - minY + 3, width + 1).BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(80, 80, 80)
            map.Cells(4, origin + 1).Resize(1, width).Font.Size = 20
            map.Cells(5, origin).Resize(maxY - minY + 1, 1).Font.Size = 20
            For col = minX To maxX: map.Cells(4, origin + 1 + col - minX).Value2 = col: Next col
            For row = minY To maxY: map.Cells(5 + row - minY, origin).Value2 = row: Next row
        Next phase
        For Each item In ordered
            xy = Split(CStr(item), "|")
            For phase = 1 To phaseCount
                result = MapResult(counts, errors, "WT" & phase, CStr(detail(0)), CStr(item), hard, wanted, includedMask, priorPatterns)
                chipCat = 0
                If categoryByWT(phase).Exists(CStr(item)) Then chipCat = CLng(categoryByWT(phase)(CStr(item)))
                If chipCat = 1 Or chipCat = 2 Then
                    If Not evidence.Exists("WT" & phase & "|" & CStr(item)) Then result = 0
                End If
                n = n + 1: output(n, 1) = name: output(n, 2) = CLng(detail(0)): output(n, 3) = detail(1): output(n, 4) = detail(2)
                output(n, 5) = "WT" & phase: output(n, 6) = CLng(xy(0)): output(n, 7) = CLng(xy(1)): output(n, 9) = MapStatus(result)
                If VarType(result) <> vbString Then
                    output(n, 8) = result
                    If CLng(result) > maximumCount Then maximumCount = CLng(result)
                End If
                origin = 1 + (phase - 1) * (width + 2)
                With map.Cells(5 + CLng(xy(1)) - minY, origin + 1 + CLng(xy(0)) - minX)
                    If VarType(result) = vbString Then .ClearContents Else .Value2 = CLng(result)
                    .Font.Size = 24: .Font.Color = vbBlack: .NumberFormat = "0;-0;;@"
                    If VarType(result) = vbString Then .Interior.Color = MapColor(result) Else .Interior.Color = vbWhite
                    .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(150, 150, 150)
                End With
            Next phase
        Next item
        ApplyMapColorScale map, minX, maxX, minY, maxY, phaseCount, maximumCount
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
    BLK_ApplyPresentation wb
    Exit Function
Failed:
    BLK_BuildFailMaps = "Map build failed (" & Err.Number & "): " & Err.Description
End Function

Public Sub BLK_InvalidateCoverageMap(ByVal wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("FailBLK Loss Map")
    On Error GoTo 0
    If Not ws Is Nothing Then ws.UsedRange.UnMerge: ws.UsedRange.Clear
End Sub

Public Sub BLK_BuildCoverageMap(ByVal wb As Workbook)
    'Right cell counts BLKs with no Address anywhere in the Pattern interval; legacy reports retain old shortfall behavior.
    Dim ws As Worksheet, source As Worksheet, report As Worksheet
    Dim summaries As Variant, data As Variant, names As Object, coords As Object, losses As Object, seen As Object
    Dim cats(1 To 3) As Object, ordered As Collection, item As Variant, name As Variant, xy As Variant, evidence As Object
    Dim i As Long, j As Long, phase As Long, phaseCount As Long, last As Long, first As Long, take As Long
    Dim minX As Long, maxX As Long, minY As Long, maxY As Long, x As Long, y As Long
    Dim top As Long, height As Long, panelWidth As Long, origin As Long, row As Long, col As Long
    Dim key As String, blockKey As String, chip As String, status As String, cat As Long, loss As Long
    Dim fail As Long, maximum As Long, cells As Range, colorRule As Object, relaxedCoverage As Boolean
    RevealMapSources wb, True
    Set source = wb.Worksheets("BLK Fail Count"): last = source.Cells(source.Rows.Count, 1).End(xlUp).Row
    If last < 2 Then Err.Raise vbObjectError + 492, , "Build Fail Maps before the Address Map."
    summaries = source.Range("A2:I" & last).Value2
    Set names = CreateObject("Scripting.Dictionary"): Set coords = CreateObject("Scripting.Dictionary")
    Set losses = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary")
    phaseCount = BLK_PhaseCount(wb)
    Set evidence = ChipFailEvidence(wb)
    For phase = 1 To phaseCount: Set cats(phase) = BLK_Categories(wb, "WT1"): Next phase
    For i = 1 To UBound(summaries, 1)
        names(CStr(summaries(i, 1))) = True
        coords(CStr(summaries(i, 6)) & "|" & CStr(summaries(i, 7))) = True
    Next i
    Set ordered = SortedDUTs(coords)
    xy = Split(CStr(ordered(1)), "|"): minX = CLng(xy(0)): maxX = minX: minY = CLng(xy(1)): maxY = minY
    For Each item In ordered
        xy = Split(CStr(item), "|"): x = CLng(xy(0)): y = CLng(xy(1))
        If x < minX Then minX = x
        If x > maxX Then maxX = x
        If y < minY Then minY = y
        If y > maxY Then maxY = y
    Next item
    Set report = wb.Worksheets("Missing BLK Address"): last = report.Cells(report.Rows.Count, 1).End(xlUp).Row
    relaxedCoverage = (CStr(report.Range("M4").Value2) = "Pattern Address Count")
    For first = 5 To last Step 2048
        take = last - first + 1: If take > 2048 Then take = 2048
        data = report.Cells(first, 1).Resize(take, 12).Value2
        For i = 1 To UBound(data, 1)
            If (relaxedCoverage And CStr(data(i, 12)) = "No Address") Or (Not relaxedCoverage And CDbl(data(i, 11)) > 0) Then
                key = CStr(data(i, 1)) & "|" & CStr(data(i, 5)) & "|" & CStr(data(i, 6)) & "|" & CStr(data(i, 7))
                blockKey = key & "|" & CStr(data(i, 8))
                If Not seen.Exists(blockKey) Then
                    seen.Add blockKey, True
                    If Not losses.Exists(key) Then losses.Add key, 0
                    losses(key) = CLng(losses(key)) + 1
                End If
            End If
        Next i
    Next first
    On Error Resume Next
    Set ws = wb.Worksheets("FailBLK Loss Map")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=source): ws.Name = "FailBLK Loss Map"
        ws.CustomProperties.Add Name:="ASCCoverageMap", Value:="1"
    End If
    ws.UsedRange.UnMerge: ws.UsedRange.Clear
    panelWidth = (maxX - minX + 1) * 2 + 2: height = maxY - minY + 7: top = 1
    If CDbl(height) * names.Count > ws.Rows.Count Then Err.Raise vbObjectError + 492, , "Address Map exceeds Excel row limit."
    With ws.Cells(1, 1).Resize(height * names.Count, panelWidth * phaseCount)
        .Font.Name = "Calibri": .Font.Size = 24: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ColumnWidth = 13: .RowHeight = 40: .NumberFormat = "0;-0;;@"
    End With
    For Each name In names
        ws.Cells(top, 1).Resize(1, panelWidth * phaseCount).Merge
        With ws.Cells(top, 1)
            .Value2 = CStr(name) & " | Left: Fail BLK / Right: Loss BLK"
            .Font.Bold = True: .Font.Color = vbWhite: .Interior.Color = RGB(23, 50, 77)
        End With
        maximum = 2
        For phase = 1 To phaseCount
            origin = 1 + (phase - 1) * panelWidth
            ws.Cells(top + 1, origin + 1).Resize(1, panelWidth - 2).Merge
            ws.Cells(top + 1, origin + 1).Value2 = "WT" & phase
            ws.Cells(top + 1, origin + 1).Font.Bold = True
            For x = minX To maxX
                col = origin + 1 + (x - minX) * 2
                ws.Cells(top + 2, col).Resize(1, 2).Merge
                ws.Cells(top + 2, col).Value2 = x: ws.Cells(top + 2, col).Font.Size = 20
                ws.Cells(top + 3, col).Value2 = "Fail": ws.Cells(top + 3, col + 1).Value2 = "Loss"
                ws.Cells(top + 2, col).Resize(2, 2).Borders.LineStyle = xlContinuous
            Next x
            For y = minY To maxY
                ws.Cells(top + 4 + y - minY, origin).Value2 = y
                ws.Cells(top + 4 + y - minY, origin).Font.Size = 20
                ws.Cells(top + 4 + y - minY, origin).Borders.LineStyle = xlContinuous
            Next y
            ws.Cells(top + 1, origin).Resize(maxY - minY + 5, panelWidth - 1).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        Next phase
        For i = 1 To UBound(summaries, 1)
            If CStr(summaries(i, 1)) = CStr(name) Then
                phase = CLng(Mid$(CStr(summaries(i, 5)), 3))
                x = CLng(summaries(i, 6)): y = CLng(summaries(i, 7)): chip = x & "|" & y
                origin = 1 + (phase - 1) * panelWidth
                row = top + 4 + y - minY: col = origin + 1 + (x - minX) * 2
                Set cells = ws.Cells(row, col).Resize(1, 2)
                cells.Borders.LineStyle = xlContinuous: cells.Borders.Color = RGB(150, 150, 150)
                cat = 0: If cats(phase).Exists(chip) Then cat = CLng(cats(phase)(chip))
                If (cat <> 1 And cat <> 2) Or evidence.Exists("WT" & phase & "|" & chip) Then
                    status = CStr(summaries(i, 9))
                    If status = "Pass" Or status = "Fail" Then
                        fail = CLng(summaries(i, 8)): key = "WT" & phase & "|" & CStr(name) & "|" & chip
                        loss = 0: If losses.Exists(key) Then loss = CLng(losses(key))
                        If loss > fail Then Err.Raise vbObjectError + 492, , "Loss BLKs exceed Map Fail BLKs: " & key
                        ws.Cells(row, col).Value2 = fail: ws.Cells(row, col + 1).Value2 = loss
                        If fail > maximum Then maximum = fail
                    Else
                        cells.Interior.Color = MapColor(status)
                    End If
                End If
            End If
        Next i
        For phase = 1 To phaseCount
            origin = 1 + (phase - 1) * panelWidth
            Set cells = ws.Cells(top + 4, origin + 1).Resize(maxY - minY + 1, panelWidth - 2)
            Set colorRule = cells.FormatConditions.AddColorScale(ColorScaleType:=2)
            With colorRule.ColorScaleCriteria(1)
                .Type = xlConditionValueNumber: .Value = 1: .FormatColor.Color = RGB(255, 235, 132)
            End With
            With colorRule.ColorScaleCriteria(2)
                .Type = xlConditionValueNumber: .Value = maximum: .FormatColor.Color = RGB(248, 105, 107)
            End With
            Set colorRule = cells.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(ISNUMBER(" & cells.Cells(1, 1).Address(False, False) & ")," & cells.Cells(1, 1).Address(False, False) & "=0)")
            colorRule.SetFirstPriority: colorRule.StopIfTrue = True
            colorRule.Interior.Color = vbWhite: colorRule.Font.Color = vbWhite
        Next phase
        top = top + height
    Next name
    ws.Activate: Application.ActiveWindow.DisplayGridlines = False
    BLK_OrderSheets wb
End Sub

Public Sub BLK_ClearChipMonitorMap(ByVal wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("W&R Map")
    If Not ws Is Nothing Then ws.UsedRange.UnMerge: ws.UsedRange.Clear
End Sub

Private Function ConfirmedFunctionPassChips(ByVal wb As Workbook, ByVal coords As Object, ByVal phaseCount As Long) As Object
    Dim wanted As Object, counts As Object, addressFail As Object, result As Object, cfg As Variant
    Dim chip As Variant, tno As Variant, i As Long, phase As Long, key As String, recordKey As String
    Dim allZero As Boolean, zeroMask As String
    Set wanted = CreateObject("Scripting.Dictionary"): Set result = CreateObject("Scripting.Dictionary")
    cfg = wb.Worksheets("Config").Range("D2:D901").Value2
    For i = 1 To UBound(cfg, 1)
        If Len(Trim$(CStr(cfg(i, 1)))) > 0 Then wanted(CStr(CLng(cfg(i, 1)))) = True
    Next i
    If wanted.Count = 0 Then Set ConfirmedFunctionPassChips = result: Exit Function
    Set counts = LoadMapCounts(wb.Worksheets("FBC by BLK"), wanted)
    Set addressFail = ImportFailAddresses.MapChipAddressEvidence(wb, True)
    zeroMask = String$(256, "0")
    For phase = 1 To phaseCount
        If CStr(wb.Worksheets("Chip Locations").Cells(phase + 1, 5).Value2) <> "Data Error" Then
            For Each chip In coords
                key = "WT" & phase & "|" & CStr(chip)
                If Not addressFail.Exists(key) Then
                    allZero = True
                    For Each tno In wanted
                        recordKey = "WT" & phase & "|" & CStr(tno) & "|" & CStr(chip)
                        If Not counts.Exists(recordKey) Then
                            allZero = False: Exit For
                        End If
                        If CStr(counts(recordKey)) <> zeroMask Then allZero = False: Exit For
                    Next tno
                    If allZero Then result.Add key, True
                End If
            Next chip
        End If
    Next phase
    Set ConfirmedFunctionPassChips = result
End Function

Public Sub BLK_BuildChipMonitorMap(ByVal wb As Workbook, ByVal outputWorkbook As Workbook)
    'Per Chip counts: upper-left Read, upper-right Write, lower-left Both, lower-right None.
    'Presence is aggregated across all classified Function Patterns within each WT.
    Dim ws As Worksheet, locations As Worksheet, summary As Worksheet
    Dim passChips As Object, chipCounts As Object, seenAddress As Object, detail As Worksheet
    Dim totals As Variant, category As Long, addressKey As String, categoryName As String, bitMask As Long, maxCount As Long
    Dim coords As Object, masks As Object, ordered As Collection, item As Variant, xy As Variant, data As Variant
    Dim phaseCount As Long, phase As Long, last As Long, first As Long, take As Long, i As Long, j As Long
    Dim x As Long, y As Long, minX As Long, maxX As Long, minY As Long, maxY As Long
    Dim panelWidth As Long, origin As Long, row As Long, col As Long, mask As Long, addBits As Variant
    Dim oldLastCol As Long, oldLastRow As Long, newLastRow As Long
    Dim key As String, wt As String, chip As String, value As Variant, count As Double
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    RevealMapSources outputWorkbook, True
    Set summary = outputWorkbook.Worksheets("Monitor BLK Summary")
    If CStr(summary.Range("J4").Value2) <> "Write Fail" Or CStr(summary.Range("M4").Value2) <> "No monitor item Fail" Then Err.Raise vbObjectError + 496, , "Refresh Monitor analysis before building Chip Map."
    If Left$(CStr(summary.Range("A2").Value2), 14) = "Source updated" Then Err.Raise vbObjectError + 496, , "Monitor analysis is stale; refresh it first."
    Set locations = wb.Worksheets("Chip Locations"): last = locations.Cells(locations.Rows.Count, 1).End(xlUp).Row
    If last < 2 Then Err.Raise vbObjectError + 496, , "Import Raw Chip locations before building Monitor Map."
    Set coords = CreateObject("Scripting.Dictionary")
    data = locations.Range("A2:B" & last).Value2
    For i = 1 To UBound(data, 1)
        coords(Coordinate(CStr(data(i, 1))) & "|" & Coordinate(CStr(data(i, 2)))) = True
    Next i
    Set ordered = SortedDUTs(coords): xy = Split(CStr(ordered(1)), "|")
    minX = CLng(xy(0)): maxX = minX: minY = CLng(xy(1)): maxY = minY
    For Each item In ordered
        xy = Split(CStr(item), "|"): x = CLng(xy(0)): y = CLng(xy(1))
        If x < minX Then minX = x
        If x > maxX Then maxX = x
        If y < minY Then minY = y
        If y > maxY Then maxY = y
    Next item
    If CDbl(maxX) - minX > 99 Or CDbl(maxY) - minY > 99 Then Err.Raise vbObjectError + 496, , "Monitor Map coordinate span exceeds 100."
    phaseCount = BLK_PhaseCount(wb)
    Set passChips = ConfirmedFunctionPassChips(wb, coords, phaseCount)
    Set masks = CreateObject("Scripting.Dictionary"): addBits = Array(2, 1, 4, 8)
    last = summary.Cells(summary.Rows.Count, 1).End(xlUp).Row
    For first = 5 To last Step 512
        take = last - first + 1: If take > 512 Then take = 512
        data = summary.Cells(first, 1).Resize(take, 16).Value2
        For i = 1 To take
            wt = CStr(data(i, 1))
            If wt <> "WT1" And wt <> "WT2" And wt <> "WT3" Then Err.Raise vbObjectError + 496, , "Invalid WT in Monitor summary."
            If CLng(Right$(wt, 1)) > phaseCount Then Err.Raise vbObjectError + 496, , "Monitor summary contains an unimported WT."
            chip = Coordinate(CStr(data(i, 4))) & "|" & Coordinate(CStr(data(i, 5)))
            If Not coords.Exists(chip) Then Err.Raise vbObjectError + 496, , "Monitor Chip is absent from Raw locations."
            key = wt & "|" & chip: mask = 0
            If masks.Exists(key) Then mask = CLng(masks(key))
            count = 0
            For j = 10 To 13
                value = data(i, j)
                If IsError(value) Then Err.Raise vbObjectError + 496, , "Invalid Monitor classification count."
                If Len(CStr(value)) > 0 Then
                    If Not IsNumeric(value) Then Err.Raise vbObjectError + 496, , "Monitor classification count must be integer."
                    If CDbl(value) < 0 Or CDbl(value) <> Fix(CDbl(value)) Then Err.Raise vbObjectError + 496, , "Monitor classification count must be nonnegative integer."
                    count = count + CDbl(value)
                    If CDbl(value) > 0 Then mask = mask Or CLng(addBits(j - 10))
                End If
            Next j
            If count <> CDbl(data(i, 9)) Then Err.Raise vbObjectError + 496, , "Monitor category counts do not match observed Function Addresses."
            If mask > 0 Then masks(key) = mask
        Next i
    Next first
    Set chipCounts = CreateObject("Scripting.Dictionary"): Set seenAddress = CreateObject("Scripting.Dictionary")
    Set detail = outputWorkbook.Worksheets("Monitor Fail Analysis")
    last = detail.Cells(detail.Rows.Count, 1).End(xlUp).Row
    For first = 5 To last Step 1024
        take = last - first + 1: If take > 1024 Then take = 1024
        data = detail.Cells(first, 1).Resize(take, 16).Value2
        For i = 1 To take
            wt = CStr(data(i, 1)): chip = CStr(data(i, 4)) & "|" & CStr(data(i, 5)): key = wt & "|" & chip
            If Not coords.Exists(chip) Then Err.Raise vbObjectError + 496, , "Monitor Address Chip is absent from Raw locations."
            categoryName = CStr(data(i, 15))
            Select Case categoryName
                Case "Read Fail": category = 0: bitMask = 1
                Case "Write Fail": category = 1: bitMask = 2
                Case "Write&Read Fail": category = 2: bitMask = 4
                Case "No monitor item Fail": category = 3: bitMask = 8
                Case Else: Err.Raise vbObjectError + 496, , "Unknown Monitor classification."
            End Select
            If Not masks.Exists(key) Then Err.Raise vbObjectError + 496, , "Monitor detail does not match BLK analysis."
            If (CLng(masks(key)) And bitMask) = 0 Then Err.Raise vbObjectError + 496, , "Monitor category differs from BLK analysis."
            addressKey = key
            For j = 6 To 9: addressKey = addressKey & "|" & CStr(data(i, j)): Next j
            If seenAddress.Exists(addressKey) Then
                If CStr(seenAddress(addressKey)) <> categoryName Then Err.Raise vbObjectError + 496, , "Conflicting Monitor category for one physical Address."
            Else
                seenAddress.Add addressKey, categoryName
                If chipCounts.Exists(key) Then totals = chipCounts(key) Else totals = Array(0&, 0&, 0&, 0&)
                totals(category) = CLng(totals(category)) + 1: chipCounts(key) = totals
                If CLng(totals(category)) > maxCount Then maxCount = CLng(totals(category))
            End If
        Next i
    Next first
    On Error Resume Next
    Set ws = outputWorkbook.Worksheets("W&R Map")
    On Error GoTo Failed
    If ws Is Nothing Then
        Set ws = outputWorkbook.Worksheets.Add(After:=summary): ws.Name = "W&R Map"
    End If
    oldLastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    oldLastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    ws.UsedRange.UnMerge: ws.UsedRange.Clear
    panelWidth = 2 * (maxX - minX + 1) + 2
    newLastRow = 2 * (maxY - minY + 1) + 5
    If oldLastCol > panelWidth * phaseCount Then ws.Range(ws.Columns(panelWidth * phaseCount + 1), ws.Columns(oldLastCol)).Delete
    If oldLastRow > newLastRow Then ws.Range(ws.Rows(newLastRow + 1), ws.Rows(oldLastRow)).Delete
    With ws.Cells(1, 1).Resize(newLastRow, panelWidth * phaseCount)
        .Font.Name = "Calibri": .Font.Size = 24: .Font.Color = vbBlack
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ColumnWidth = 8.5: .RowHeight = 42: .Interior.Pattern = xlNone
        If maxCount >= 100 Then .ColumnWidth = 11
        If maxCount >= 1000 Then .ColumnWidth = 14
    End With
    ws.Cells(1, 1).Resize(1, panelWidth * phaseCount).Merge
    With ws.Cells(1, 1)
        .Value2 = "W&R Map"
        .WrapText = True: .RowHeight = 70
        .Font.Bold = True: .Font.Color = vbWhite: .Interior.Color = RGB(23, 50, 77)
    End With
    ws.Rows(2).RowHeight = 12
    For phase = 1 To phaseCount
        origin = 1 + (phase - 1) * panelWidth
        ws.Cells(3, origin + 1).Resize(1, panelWidth - 2).Merge
        ws.Cells(3, origin + 1).Value2 = "WT" & phase: ws.Cells(3, origin + 1).Font.Bold = True
        ws.Cells(3, origin + 1).Resize(1, panelWidth - 2).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        For x = minX To maxX
            col = origin + 1 + 2 * (x - minX)
            With ws.Cells(4, col).Resize(1, 2)
                .Merge: .Cells(1, 1).Value2 = x: .Font.Size = 20: .Borders.LineStyle = xlContinuous
            End With
        Next x
        For y = minY To maxY
            row = 5 + 2 * (y - minY)
            With ws.Cells(row, origin).Resize(2, 1)
                .Merge: .Cells(1, 1).Value2 = y: .Font.Size = 20: .Borders.LineStyle = xlContinuous
            End With
        Next y
        For Each item In ordered
            xy = Split(CStr(item), "|")
            row = 5 + 2 * (CLng(xy(1)) - minY): col = origin + 1 + 2 * (CLng(xy(0)) - minX)
            With ws.Cells(row, col).Resize(2, 2)
                .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(150, 150, 150)
                .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, Color:=RGB(80, 80, 80)
            End With
            key = "WT" & phase & "|" & CStr(item): mask = 0
            If masks.Exists(key) Then mask = CLng(masks(key))
            totals = Array(0&, 0&, 0&, 0&)
            If chipCounts.Exists(key) Then totals = chipCounts(key)
            If mask = 0 And passChips.Exists(key) Then ws.Cells(row, col).Resize(2, 2).Interior.Color = RGB(0, 176, 240)
            If CLng(totals(0)) > 0 Then
                With ws.Cells(row, col): .Value2 = CLng(totals(0)): .Interior.Pattern = xlNone: End With
            End If
            If CLng(totals(1)) > 0 Then
                With ws.Cells(row, col + 1): .Value2 = CLng(totals(1)): .Interior.Color = RGB(255, 217, 202): End With
            End If
            If CLng(totals(2)) > 0 Then
                With ws.Cells(row + 1, col): .Value2 = CLng(totals(2)): .Interior.Color = RGB(255, 102, 0): End With
            End If
            If CLng(totals(3)) > 0 Then
                With ws.Cells(row + 1, col + 1): .Value2 = CLng(totals(3)): .Font.Color = RGB(255, 255, 255): .Interior.Pattern = xlNone: End With
            End If
        Next item
        ws.Cells(3, origin).Resize(2 * (maxY - minY + 1) + 2, panelWidth - 1).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    Next phase
    ws.Activate: Application.ActiveWindow.DisplayGridlines = False
    BLK_OrderSheets outputWorkbook
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    BLK_ClearChipMonitorMap outputWorkbook
    Err.Raise errorNumber, , errorText
End Sub

Public Sub BLK_MigrateReportNames(ByVal wb As Workbook)
    Dim old As Worksheet, current As Worksheet, pair As Variant
    For Each pair In Array(Array("Chip Monitor vs Function Map", "W&R Map"), Array("Fail BLK Address Map", "FailBLK Loss Map"))
        Set old = Nothing: Set current = Nothing
        On Error Resume Next
        Set old = wb.Worksheets(CStr(pair(0))): Set current = wb.Worksheets(CStr(pair(1)))
        On Error GoTo 0
        If current Is Nothing Then
            If Not old Is Nothing Then old.Name = CStr(pair(1))
        End If
    Next pair
End Sub

Private Function ReportSheet(ByVal wb As Workbook, ByVal name As String) As Worksheet
    Dim ws As Worksheet, i As Long
    On Error Resume Next
    Set ws = wb.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)): ws.Name = name
    ws.Visible = xlSheetVisible
    If name = "Chip Fail Details" Then
        For i = ws.ChartObjects.Count To 1 Step -1
            If ws.ChartObjects(i).Name = "blkVoltageChart" Then ws.ChartObjects(i).Delete
        Next i
    End If
    For i = ws.ListObjects.Count To 1 Step -1: ws.ListObjects(i).Delete: Next i
    ws.UsedRange.UnMerge: ws.UsedRange.Clear
    Set ReportSheet = ws
End Function

Private Function AppendReportTable(ByVal source As Worksheet, ByVal header As Long, ByVal width As Long, ByVal dest As Worksheet, ByVal top As Long, ByVal title As String, ByVal tableName As String) As Long
    Dim last As Long, start As Long, take As Long, n As Long, target As Long, table As ListObject
    last = source.Cells(source.Rows.Count, 1).End(xlUp).Row
    If last < header Then last = header
    n = last - header
    If top + n + 2 > dest.Rows.Count Then Err.Raise vbObjectError + 510, , "Combined Fail Analysis exceeds Excel row limit."
    With dest.Cells(top, 1).Resize(1, width)
        .Merge: .Cells(1, 1).Value2 = title & " | rows=" & n
        .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: .RowHeight = 28
    End With
    If tableName = "tblMonitorAddr" Then dest.Columns("J:L").NumberFormat = "@"
    target = top + 1
    For start = header To last Step 512
        take = last - start + 1: If take > 512 Then take = 512
        dest.Cells(target, 1).Resize(take, width).Value2 = source.Cells(start, 1).Resize(take, width).Value2
        target = target + take
    Next start
    Set table = dest.ListObjects.Add(SourceType:=xlSrcRange, Source:=dest.Cells(top + 1, 1).Resize(n + 1, width), XlListObjectHasHeaders:=xlYes)
    table.Name = tableName: table.TableStyle = "TableStyleMedium2"
    With table.Range
        .Font.Name = "Calibri": .Font.Size = 11: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    With table.HeaderRowRange: .Orientation = 90: .WrapText = True: .RowHeight = 130: End With
    If n > 0 Then
        table.DataBodyRange.RowHeight = 26
        If tableName = "tblMonitorAddr" Then
            table.DataBodyRange.RowHeight = 52: table.ListColumns(10).DataBodyRange.WrapText = True
            table.ListColumns(11).DataBodyRange.WrapText = True: table.ListColumns(12).DataBodyRange.WrapText = True
            With table.ListColumns(16).DataBodyRange.Font: .Color = RGB(0, 102, 204): .Underline = xlUnderlineStyleSingle: End With
        End If
    End If
    AppendReportTable = target + 2
End Function

Public Sub BLK_BuildIntegratedAnalysis(ByVal wb As Workbook)
    Dim ws As Worksheet, top As Long, firstBLK As Long, firstAddress As Long, name As Variant, required As Double
    required = wb.Worksheets("BLK Fail Count").Cells(wb.Worksheets("BLK Fail Count").Rows.Count, 1).End(xlUp).Row
    required = required + wb.Worksheets("Monitor BLK Summary").Cells(wb.Worksheets("Monitor BLK Summary").Rows.Count, 1).End(xlUp).Row
    required = required + wb.Worksheets("Monitor Fail Analysis").Cells(wb.Worksheets("Monitor Fail Analysis").Rows.Count, 1).End(xlUp).Row + 20
    If required > 1048576 Then Err.Raise vbObjectError + 510, , "Combined Fail Analysis exceeds Excel row limit; source outputs retained."
    Set ws = ReportSheet(wb, "Fail Analysis")
    ws.Range("A1:P1").Merge: ws.Range("A1").Value2 = "Fail Analysis | Chip > BLK > Address"
    ws.Range("A3:P3").Merge: ws.Range("A3").Value2 = CStr(wb.Worksheets("Monitor Fail Analysis").Range("A2").Value2)
    ws.Columns("A:P").ColumnWidth = 12: ws.Columns("B").ColumnWidth = 18: ws.Columns("J:L").ColumnWidth = 24: ws.Columns("O:P").ColumnWidth = 28
    top = AppendReportTable(wb.Worksheets("BLK Fail Count"), 1, 9, ws, 5, "Chip Fail BLK Count", "tblChipFail")
    firstBLK = top: top = AppendReportTable(wb.Worksheets("Monitor BLK Summary"), 4, 16, ws, top, "BLK Monitor Summary", "tblBLKMonitor")
    firstAddress = top: top = AppendReportTable(wb.Worksheets("Monitor Fail Analysis"), 4, 16, ws, top, "Function / Monitor Address Detail", "tblMonitorAddr")
    ws.Range("A2:D2").Merge: ws.Range("F2:I2").Merge: ws.Range("K2:N2").Merge
    ws.Hyperlinks.Add ws.Range("A2"), "", "'Fail Analysis'!A5", , "Chip summary"
    ws.Hyperlinks.Add ws.Range("F2"), "", "'Fail Analysis'!A" & firstBLK, , "BLK summary"
    ws.Hyperlinks.Add ws.Range("K2"), "", "'Fail Analysis'!A" & firstAddress, , "Address detail"
    With ws.Range("A1:P3"): .Font.Name = "Calibri": .Font.Size = 11: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 28: End With
    ws.Range("A3:P3").WrapText = True: ws.Rows(3).RowHeight = 42
    With ws.UsedRange: .Font.Name = "Calibri": .Font.Size = 11: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: End With
    ws.Activate: Application.ActiveWindow.DisplayGridlines = False
    For Each name In Array("BLK Fail Count", "Monitor BLK Summary", "Monitor Fail Analysis"): wb.Worksheets(CStr(name)).Visible = xlSheetVeryHidden: Next name
End Sub

Public Sub BLK_BuildImportInfo(ByVal wb As Workbook)
    Dim ws As Worksheet, idx As Worksheet, log As Worksheet, loc As Worksheet, data As Variant, routes(1 To 3, 1 To 8) As Variant
    Dim p As Long, a As Long, f As Long, selected As Long, last As Long, messageRow As Long
    Set ws = ReportSheet(wb, "Import Info"): Set log = wb.Worksheets("Log"): Set loc = wb.Worksheets("Chip Locations")
    ws.Range("A1:H1").Merge: ws.Range("A1").Value2 = "Import Info | Raw Data Routes / Log / Chip Locations"
    ws.Range("A3:H3").Value2 = Array("WT", "Raw Data Route", "LOT", "WNO", "STEP", "Imported At", "FBC Raw Route", "Address Raw Route")
    On Error Resume Next
    Set idx = wb.Worksheets("Import Index")
    On Error GoTo 0
    If Not idx Is Nothing Then
        data = idx.Range("H2:O7").Value2
        For p = 1 To 3
            a = p: f = p + 3: selected = a
            If Val(CStr(data(f, 8))) > Val(CStr(data(a, 8))) Or Len(CStr(data(a, 7))) = 0 Then selected = f
            routes(p, 1) = "WT" & p: routes(p, 2) = data(selected, 7): routes(p, 3) = data(selected, 3)
            routes(p, 4) = data(selected, 4): routes(p, 5) = data(selected, 5): routes(p, 6) = data(selected, 8)
            routes(p, 7) = data(f, 7): routes(p, 8) = data(a, 7)
        Next p
    End If
    ws.Range("A4:H6").Value2 = routes: ws.Range("D4:D6").NumberFormat = "00": ws.Range("F4:F6").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Range("A8:D8").Merge: ws.Range("A8").Value2 = "Latest operation messages"
    ws.Range("A9:D13").Value2 = log.Range("A1:D5").Value2: ws.Range("A10:A13").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Range("A16:B16").Merge: ws.Range("A16").Value2 = "Chip Locations"
    last = loc.Cells(loc.Rows.Count, 1).End(xlUp).Row: If last < 1 Then last = 1
    ws.Range("A17").Resize(last, 2).Value2 = loc.Range("A1").Resize(last, 2).Value2
    ws.Range("D17:F20").Value2 = loc.Range("D1:F4").Value2
    ws.Range("H16").Value2 = "Location message": ws.Range("H17").Value2 = loc.Range("H1").Value2
    With ws.UsedRange
        .Font.Name = "Calibri": .Font.Size = 11: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 26
    End With
    ws.Columns("A").ColumnWidth = 24: ws.Columns("B").ColumnWidth = 70: ws.Columns("C:F").ColumnWidth = 24: ws.Columns("G:H").ColumnWidth = 60
    ws.Range("B4:B6").WrapText = True: ws.Range("G4:H6").WrapText = True: ws.Rows("4:6").RowHeight = 52
    For messageRow = 9 To 13: ws.Cells(messageRow, 4).Resize(1, 5).Merge: Next messageRow
    ws.Range("D9:H13").WrapText = True: ws.Rows("10:13").RowHeight = 65
    With ws.Range("A1:H1"): .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: End With
    For Each data In Array("A3:H3", "A8:D9", "A16:B17", "D17:F17")
        With ws.Range(CStr(data)): .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: .Borders.LineStyle = xlContinuous: End With
    Next data
    wb.Worksheets("Config").Activate
    log.Visible = xlSheetVeryHidden: loc.Visible = xlSheetVeryHidden
    BLK_OrderSheets wb
End Sub

Public Sub BLK_PublishReports(ByVal wb As Workbook)
    BLK_MigrateReportNames wb
    RevealMapSources wb, True
    BLK_StandardizeMaps wb
    BLK_BuildIntegratedAnalysis wb
    BLK_BuildImportInfo wb
    BLK_BuildUnifiedMaps wb
End Sub

Private Sub MapFrame(ByVal ws As Worksheet, ByVal top As Long, ByVal title As String, ByVal minX As Long, ByVal maxX As Long, ByVal minY As Long, ByVal maxY As Long, ByVal phases As Long)
    Dim span As Long, phase As Long, origin As Long, x As Long, y As Long, c As Long, r As Long
    span = 2 * (maxX - minX + 1) + 2
    With ws.Cells(top, 1).Resize(2 * (maxY - minY + 1) + 6, span * phases)
        .Font.Name = "Calibri": .Font.Size = 24: .Font.Color = vbBlack
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .ColumnWidth = 6.5: .RowHeight = 36
    End With
    ws.Cells(top, 1).Resize(1, span * phases).Merge
    With ws.Cells(top, 1): .Value2 = title: .WrapText = True: .RowHeight = 60: .Font.Bold = True: .Font.Color = vbWhite: .Interior.Color = RGB(23, 50, 77): End With
    ws.Rows(top + 1).RowHeight = 12
    For phase = 1 To phases
        origin = 1 + (phase - 1) * span
        With ws.Cells(top + 2, origin + 1).Resize(1, span - 2)
            .Merge: .Cells(1, 1).Value2 = "WT" & phase: .Font.Bold = True: .BorderAround LineStyle:=xlContinuous, Weight:=xlThin
        End With
        For x = minX To maxX
            c = origin + 1 + 2 * (x - minX)
            With ws.Cells(top + 3, c).Resize(1, 2)
                .Merge: .Cells(1, 1).Value2 = x: .NumberFormat = "0": .Font.Size = 20: .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(80, 80, 80)
            End With
        Next x
        For y = minY To maxY
            r = top + 4 + 2 * (y - minY)
            With ws.Cells(r, origin).Resize(2, 1)
                .Merge: .Cells(1, 1).Value2 = y: .NumberFormat = "0": .Font.Size = 20: .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(80, 80, 80)
            End With
        Next y
        ws.Cells(top + 2, origin).Resize(2 * (maxY - minY + 1) + 2, span - 1).BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(80, 80, 80)
    Next phase
End Sub

Private Sub ChipBorder(ByVal chip As Range)
    chip.Borders.LineStyle = xlContinuous: chip.Borders.Color = RGB(80, 80, 80): chip.Borders.Weight = xlThin
    If chip.Cells(1, 1).MergeArea.Cells.CountLarge < chip.Cells.CountLarge Then
        chip.Borders(xlInsideHorizontal).LineStyle = xlDash: chip.Borders(xlInsideHorizontal).Color = RGB(150, 150, 150)
        chip.Borders(xlInsideVertical).LineStyle = xlDash: chip.Borders(xlInsideVertical).Color = RGB(150, 150, 150)
    End If
    chip.BorderAround LineStyle:=xlContinuous, Weight:=xlMedium, Color:=RGB(50, 50, 50)
    chip.ShrinkToFit = True
End Sub

Private Sub IndexMapCell(ByVal index As Worksheet, ByRef dest As Long, ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal wt As String, ByVal pattern As String, ByVal x As Long, ByVal y As Long, ByVal metric As String, ByVal category As String, ByVal stopTno As String, ByVal value As Variant, ByVal state As String)
    Dim row(1 To 1, 1 To 12) As Variant
    row(1, 1) = ws.Name: row(1, 2) = r: row(1, 3) = c: row(1, 4) = wt: row(1, 5) = pattern
    row(1, 6) = x: row(1, 7) = y: row(1, 8) = metric: row(1, 9) = category: row(1, 10) = stopTno
    row(1, 11) = value: row(1, 12) = state
    index.Cells(dest, 33).Resize(1, 12).Value2 = row: dest = dest + 1
End Sub

Private Sub MapHeat(ByVal ws As Worksheet, ByVal top As Long, ByVal minY As Long, ByVal maxY As Long, ByVal span As Long, ByVal phases As Long, ByVal maximum As Long)
    Dim phase As Long, origin As Long, cells As Range, rule As Object
    If maximum < 2 Then maximum = 2
    For phase = 1 To phases
        origin = 1 + (phase - 1) * span
        Set cells = ws.Cells(top + 4, origin + 1).Resize(2 * (maxY - minY + 1), span - 2)
        Set rule = cells.FormatConditions.AddColorScale(ColorScaleType:=2)
        With rule.ColorScaleCriteria(1): .Type = xlConditionValueNumber: .Value = 1: .FormatColor.Color = RGB(255, 235, 132): End With
        With rule.ColorScaleCriteria(2): .Type = xlConditionValueNumber: .Value = maximum: .FormatColor.Color = RGB(248, 105, 107): End With
    Next phase
End Sub

Public Sub BLK_StandardizeMaps(ByVal wb As Workbook)
    Dim loc As Worksheet, ws As Worksheet, index As Worksheet, summary As Worksheet
    Dim coords As Object, patterns As Object, counts As Object, loss As Object, seen As Object, pass As Object, quad As Object
    Dim ordered As Collection, item As Variant, name As Variant, detail As Variant, data As Variant, xy As Variant, q As Variant, key As String, k As String, blocked As Object, wanted As Object
    Dim i As Long, j As Long, phase As Long, phases As Long, minX As Long, maxX As Long, minY As Long, maxY As Long
    Dim last As Long, span As Long, top As Long, blockHeight As Long, origin As Long, r As Long, c As Long, dest As Long, maximum As Long, x As Long, y As Long
    Dim value As Variant, state As String, fail As Long, lost As Long, category As Variant, oldQuad As Worksheet, oldLastCol As Long, oldLastRow As Long
    Set loc = wb.Worksheets("Chip Locations"): last = loc.Cells(loc.Rows.Count, 1).End(xlUp).Row
    If last < 2 Then Exit Sub
    Set coords = CreateObject("Scripting.Dictionary"): data = loc.Range("A2:B" & last).Value2
    For i = 1 To UBound(data, 1): coords(CStr(data(i, 1)) & "|" & CStr(data(i, 2))) = True: Next i
    Set ordered = SortedDUTs(coords): xy = Split(CStr(ordered(1)), "|"): minX = CLng(xy(0)): maxX = minX: minY = CLng(xy(1)): maxY = minY
    For Each item In ordered
        xy = Split(CStr(item), "|"): x = CLng(xy(0)): y = CLng(xy(1))
        If x < minX Then minX = x
        If x > maxX Then maxX = x
        If y < minY Then minY = y
        If y > maxY Then maxY = y
    Next item
    phases = BLK_PhaseCount(wb): span = 2 * (maxX - minX + 1) + 2: blockHeight = 2 * (maxY - minY + 1) + 6
    Set pass = ConfirmedFunctionPassChips(wb, coords, phases)
    Set patterns = ReadStopPatterns(wb.Worksheets("Config"), wanted, blocked)
    Set counts = CreateObject("Scripting.Dictionary"): Set summary = wb.Worksheets("BLK Fail Count")
    last = summary.Cells(summary.Rows.Count, 1).End(xlUp).Row
    If last >= 2 Then
        data = summary.Range("A2:I" & last).Value2
        For i = 1 To UBound(data, 1): counts(CStr(data(i, 5)) & "|" & CStr(data(i, 1)) & "|" & CStr(data(i, 6)) & "|" & CStr(data(i, 7))) = Array(data(i, 8), data(i, 9)): Next i
    End If
    Set loss = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary"): Set summary = wb.Worksheets("Missing BLK Address")
    last = summary.Cells(summary.Rows.Count, 1).End(xlUp).Row
    If last >= 5 Then
        data = summary.Range("A5:O" & last).Value2
        For i = 1 To UBound(data, 1)
            If CStr(data(i, 12)) = "No Address" Then
                key = CStr(data(i, 1)) & "|" & CStr(data(i, 5)) & "|" & CStr(data(i, 6)) & "|" & CStr(data(i, 7)): k = key & "|" & CStr(data(i, 8))
                If Not seen.Exists(k) Then
                    seen(k) = True: If Not loss.Exists(key) Then loss(key) = 0
                    loss(key) = CLng(loss(key)) + 1
                End If
            End If
        Next i
    End If
    'Capture the numerical four-quadrant results before rebuilding its geometry.
    Set oldQuad = wb.Worksheets("W&R Map"): Set quad = CreateObject("Scripting.Dictionary")
    For phase = 1 To phases
        origin = 1 + (phase - 1) * span
        For Each item In ordered
            xy = Split(CStr(item), "|"): r = 5 + 2 * (CLng(xy(1)) - minY): c = origin + 1 + 2 * (CLng(xy(0)) - minX)
            quad("WT" & phase & "|" & CStr(item)) = Array(oldQuad.Cells(r, c).Value2, oldQuad.Cells(r, c + 1).Value2, oldQuad.Cells(r + 1, c).Value2, oldQuad.Cells(r + 1, c + 1).Value2)
        Next item
    Next phase
    Set index = wb.Worksheets("Import Index"): last = index.Cells(index.Rows.Count, 33).End(xlUp).Row
    index.Cells(1, 33).Resize(last, 12).ClearContents
    index.Cells(1, 33).Resize(1, 12).Value2 = Array("Sheet", "Row", "Col", "WT", "Pattern", "X", "Y", "Metric", "Category", "StopTNO", "Value", "State"): dest = 2
    Set summary = wb.Worksheets("FailBLK Loss Map"): summary.UsedRange.UnMerge: summary.UsedRange.Clear: top = 1
    For Each name In patterns
        detail = patterns(name): Set ws = wb.Worksheets(FailMapNameForPattern(wb, CStr(name)))
        ws.UsedRange.UnMerge: ws.UsedRange.Clear
        MapFrame ws, 1, ws.Name, minX, maxX, minY, maxY, phases
        MapFrame summary, top, CStr(name), minX, maxX, minY, maxY, phases
        maximum = 2
        For phase = 1 To phases
            origin = 1 + (phase - 1) * span
            For Each item In ordered
                xy = Split(CStr(item), "|"): x = CLng(xy(0)): y = CLng(xy(1)): c = origin + 1 + 2 * (x - minX): r = 5 + 2 * (y - minY)
                key = "WT" & phase & "|" & CStr(name) & "|" & CStr(item): value = Empty: state = "No Data"
                If counts.Exists(key) Then q = counts(key): value = q(0): state = CStr(q(1))
                With ws.Cells(r, c).Resize(2, 2)
                    .Merge
                    If state = "Pass" Or state = "Fail" Then
                        If Val(CStr(value)) > 0 Then .Cells(1, 1).Value2 = CLng(value)
                    Else
                        .Interior.Color = MapColor(state)
                    End If
                End With
                ChipBorder ws.Cells(r, c).Resize(2, 2)
                IndexMapCell index, dest, ws, r, c, "WT" & phase, CStr(name), x, y, "Fail BLK Count", "", CStr(detail(0)), value, state
                fail = CLng(Val(CStr(value))): lost = 0: If loss.Exists(key) Then lost = CLng(loss(key))
                If fail > maximum Then maximum = fail
                r = top + 4 + 2 * (y - minY)
                With summary.Cells(r, c).Resize(2, 2)
                    If pass.Exists("WT" & phase & "|" & CStr(item)) Then
                        .Merge: .Interior.Color = RGB(0, 176, 240)
                        IndexMapCell index, dest, summary, r, c, "WT" & phase, CStr(name), x, y, "Confirmed Pass", "", CStr(detail(0)), 0, "All Function Pass"
                    ElseIf state <> "Pass" And state <> "Fail" Then
                        .Merge: .Interior.Color = MapColor(state)
                        IndexMapCell index, dest, summary, r, c, "WT" & phase, CStr(name), x, y, "Unavailable", "", CStr(detail(0)), Empty, state
                    Else
                        summary.Cells(r, c).Resize(2, 1).Merge: summary.Cells(r, c + 1).Resize(2, 1).Merge
                        If fail > 0 Then summary.Cells(r, c).Value2 = fail
                        If lost > 0 Then summary.Cells(r, c + 1).Value2 = lost
                        IndexMapCell index, dest, summary, r, c, "WT" & phase, CStr(name), x, y, "Fail BLK Count", "", CStr(detail(0)), fail, state
                        IndexMapCell index, dest, summary, r, c + 1, "WT" & phase, CStr(name), x, y, "Loss BLK Count", "", CStr(detail(0)), lost, state
                    End If
                End With
                ChipBorder summary.Cells(r, c).Resize(2, 2)
            Next item
        Next phase
        MapHeat ws, 1, minY, maxY, span, phases, maximum
        MapHeat summary, top, minY, maxY, span, phases, maximum
        top = top + blockHeight
    Next name
    Set ws = oldQuad: ws.UsedRange.UnMerge: ws.UsedRange.Clear
    MapFrame ws, 1, "W&R Map", minX, maxX, minY, maxY, phases
    category = Array("Read Fail", "Write Fail", "Write&Read Fail", "No monitor item Fail")
    For phase = 1 To phases
        origin = 1 + (phase - 1) * span
        For Each item In ordered
            xy = Split(CStr(item), "|"): x = CLng(xy(0)): y = CLng(xy(1)): r = 5 + 2 * (y - minY): c = origin + 1 + 2 * (x - minX)
            key = "WT" & phase & "|" & CStr(item): q = quad(key)
            If pass.Exists(key) Then
                With ws.Cells(r, c).Resize(2, 2): .Merge: .Interior.Color = RGB(0, 176, 240): End With
                IndexMapCell index, dest, ws, r, c, "WT" & phase, "", x, y, "Confirmed Pass", "", "", 0, "All Function Pass"
            Else
                For j = 0 To 3
                    value = CLng(Val(CStr(q(j))))
                    With ws.Cells(r + j \ 2, c + j Mod 2)
                        If value > 0 Then
                            .Value2 = value
                            If j = 1 Then .Interior.Color = RGB(255, 217, 202)
                            If j = 2 Then .Interior.Color = RGB(255, 102, 0)
                            If j = 3 Then .Font.Color = vbWhite
                        End If
                    End With
                    IndexMapCell index, dest, ws, r + j \ 2, c + j Mod 2, "WT" & phase, "", x, y, "Monitor Address Count", CStr(category(j)), "", value, IIf(value > 0, "Recorded matches", "No recorded matches")
                Next j
            End If
            ChipBorder ws.Cells(r, c).Resize(2, 2)
        Next item
    Next phase
    For Each ws In wb.Worksheets
        If ws.Name = "W&R Map" Or ws.Name = "FailBLK Loss Map" Or Right$(ws.Name, 8) = "Fail Map" Then
            ws.Activate
            With Application.ActiveWindow
                .DisplayGridlines = False: .Zoom = IIf(phases = 3, 45, IIf(phases = 2, 65, 90))
                .FreezePanes = False: .SplitColumn = 1: .SplitRow = 4: .FreezePanes = True
                .ScrollRow = 1: .ScrollColumn = 1
            End With
        End If
    Next ws
End Sub

Private Function FailMapNameForPattern(ByVal wb As Workbook, ByVal pattern As String) As String
    Dim ws As Worksheet, i As Long
    For Each ws In wb.Worksheets
        For i = 1 To ws.CustomProperties.Count
            If ws.CustomProperties(i).Name = "ASCPattern" Then
                If CStr(ws.CustomProperties(i).Value) = pattern Then FailMapNameForPattern = ws.Name: Exit Function
            End If
        Next i
    Next ws
    Err.Raise vbObjectError + 511, , "Pattern Map missing: " & pattern
End Function

Public Function BLK_MapContext(ByVal wb As Workbook, ByVal ws As Object, ByVal target As Range) As Variant
    Dim idx As Worksheet, last As Long, data As Variant, i As Long, anchor As Range, row(0 To 11) As Variant, j As Long
    On Error GoTo Done
    If TypeName(ws) <> "Worksheet" Then Exit Function
    If ws.Name <> "Wafer Maps" And ws.Name <> "W&R Map" And ws.Name <> "FailBLK Loss Map" And Right$(ws.Name, 8) <> "Fail Map" Then Exit Function
    Set anchor = target.Cells(1, 1).MergeArea.Cells(1, 1)
    Set idx = wb.Worksheets("Import Index"): last = idx.Cells(idx.Rows.Count, 33).End(xlUp).Row
    If last < 2 Then Exit Function
    data = idx.Cells(2, 33).Resize(last - 1, 12).Value2
    For i = 1 To UBound(data, 1)
        If CStr(data(i, 1)) = ws.Name And CLng(data(i, 2)) = anchor.Row And CLng(data(i, 3)) = anchor.Column Then
            For j = 1 To 12: row(j - 1) = data(i, j): Next j
            BLK_MapContext = row: Exit Function
        End If
    Next i
Done:
End Function

Public Function BLK_IsIndexedMapCell(ByVal wb As Workbook, ByVal ws As Object, ByVal target As Range) As Boolean
    Dim info As Variant
    info = BLK_MapContext(wb, ws, target): BLK_IsIndexedMapCell = IsArray(info)
End Function

Private Function PutCollection(ByVal ws As Worksheet, ByVal top As Long, ByVal title As String, ByVal headers As Variant, ByVal rows As Collection, ByVal tableName As String) As Long
    Dim packet() As Variant, item As Variant, i As Long, j As Long, n As Long, width As Long, dest As Long, table As ListObject
    width = UBound(headers) - LBound(headers) + 1
    ws.Cells(top, 1).Resize(1, width).Merge: ws.Cells(top, 1).Value2 = title & " | rows=" & rows.Count
    ws.Cells(top, 1).Resize(1, width).Interior.Color = RGB(23, 50, 77): ws.Cells(top, 1).Font.Color = vbWhite
    For j = 0 To width - 1: ws.Cells(top + 1, j + 1).Value2 = headers(j): Next j
    dest = top + 2: i = 1
    Do While i <= rows.Count
        n = rows.Count - i + 1: If n > 512 Then n = 512
        ReDim packet(1 To n, 1 To width)
        For j = 1 To n
            item = rows(i): For width = 0 To UBound(item): packet(j, width + 1) = item(width): Next width
            i = i + 1
        Next j
        width = UBound(headers) + 1: ws.Cells(dest, 1).Resize(n, width).Value2 = packet: dest = dest + n
    Loop
    Set table = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=ws.Cells(top + 1, 1).Resize(rows.Count + 1, width), XlListObjectHasHeaders:=xlYes)
    table.Name = tableName: table.TableStyle = "TableStyleMedium2"
    table.Range.Font.Name = "Calibri": table.Range.Font.Size = 11: table.Range.HorizontalAlignment = xlCenter: table.Range.VerticalAlignment = xlCenter
    table.HeaderRowRange.Orientation = 90: table.HeaderRowRange.RowHeight = 130
    If Not table.DataBodyRange Is Nothing Then table.DataBodyRange.RowHeight = 30
    PutCollection = dest + 2
End Function

Public Sub BLK_OpenIndexedMap(ByVal wb As Workbook, ByVal map As Object, ByVal target As Range)
    Dim info As Variant, ws As Worksheet, source As Worksheet, data As Variant, item As Variant, row(0 To 15) As Variant
    Dim blocks As New Collection, addresses As New Collection, allowed As Object, loss As Object, key As String, categoryCol As Long
    Dim i As Long, j As Long, last As Long, top As Long, addrHeader As Long, t As ListObject, legacy As Boolean, r As Long
    On Error GoTo Failed
    info = BLK_MapContext(wb, map, target): If Not IsArray(info) Then Exit Sub
    If map.Name = "Wafer Maps" Then
        If InStr(1, CStr(map.Range("A2").Value2), "updated", vbTextCompare) > 0 Then Err.Raise vbObjectError + 512, , "Refresh Maps before opening updated sources."
    End If
    If InStr(1, CStr(map.Cells(1, 1).Value2), "updated", vbTextCompare) > 0 Then Err.Raise vbObjectError + 512, , "Refresh analysis before opening updated sources."
    Set allowed = CreateObject("Scripting.Dictionary"): Set loss = CreateObject("Scripting.Dictionary")
    If CStr(info(7)) = "Loss BLK Count" Then
        Set source = wb.Worksheets("Missing BLK Address"): last = source.Cells(source.Rows.Count, 1).End(xlUp).Row
        If last >= 5 Then
            data = source.Range("A5:O" & last).Value2
            For i = 1 To UBound(data, 1)
                If CStr(data(i, 1)) = CStr(info(3)) And CStr(data(i, 5)) = CStr(info(4)) And CLng(data(i, 6)) = CLng(info(5)) And CLng(data(i, 7)) = CLng(info(6)) And CStr(data(i, 12)) = "No Address" Then loss(CStr(data(i, 8))) = True
            Next i
        End If
    End If
    Select Case CStr(info(8))
        Case "Write Fail": categoryCol = 10
        Case "Read Fail": categoryCol = 11
        Case "Write&Read Fail": categoryCol = 12
        Case "No monitor item Fail": categoryCol = 13
    End Select
    Set source = wb.Worksheets("Monitor BLK Summary"): last = source.Cells(source.Rows.Count, 1).End(xlUp).Row
    If last >= 5 Then
        data = source.Range("A5:P" & last).Value2
        For i = 1 To UBound(data, 1)
            If CStr(data(i, 1)) = CStr(info(3)) And CLng(data(i, 4)) = CLng(info(5)) And CLng(data(i, 5)) = CLng(info(6)) Then
                If Len(CStr(info(4))) = 0 Or CStr(data(i, 2)) = CStr(info(4)) Then
                    legacy = True
                    If categoryCol > 0 Then If Val(CStr(data(i, categoryCol))) = 0 Then legacy = False
                    If CStr(info(7)) = "Loss BLK Count" Then If Not loss.Exists(CStr(data(i, 7))) Then legacy = False
                    If legacy Then
                        For j = 1 To 16: row(j - 1) = data(i, j): Next j
                        blocks.Add row: allowed(CStr(data(i, 2)) & "|" & CStr(data(i, 7))) = True
                    End If
                End If
            End If
        Next i
    End If
    Set source = wb.Worksheets("Monitor Fail Analysis"): last = source.Cells(source.Rows.Count, 1).End(xlUp).Row
    If last >= 5 Then
        data = source.Range("A5:P" & last).Value2
        For i = 1 To UBound(data, 1)
            If CStr(data(i, 1)) = CStr(info(3)) And CLng(data(i, 4)) = CLng(info(5)) And CLng(data(i, 5)) = CLng(info(6)) Then
                If allowed.Exists(CStr(data(i, 2)) & "|" & CStr(data(i, 6))) Then
                    If Len(CStr(info(8))) = 0 Or CStr(data(i, 15)) = CStr(info(8)) Then
                        For j = 1 To 16: row(j - 1) = data(i, j): Next j
                        addresses.Add row
                    End If
                End If
            End If
        Next i
    End If
    Set ws = ReportSheet(wb, "Chip Fail Details")
    For i = 1 To ws.CustomProperties.Count
        If ws.CustomProperties(i).Name = "ASCChipDetails" Then Exit For
    Next i
    If i > ws.CustomProperties.Count Then ws.CustomProperties.Add Name:="ASCChipDetails", Value:="1"
    ws.Columns("J:L").NumberFormat = "@"
    ws.Range("A1:P1").Merge: ws.Range("A1").Value2 = CStr(info(3)) & " | X" & info(5) & " Y" & info(6) & " | " & info(4) & " " & info(8)
    ws.Hyperlinks.Add ws.Range("A2"), "", "'" & map.Name & "'!" & map.Cells(CLng(info(1)), CLng(info(2))).Address(False, False), , "Back to Map"
    ws.Range("A3:P3").Merge: ws.Range("A3").Value2 = CStr(info(7)) & "=" & CStr(info(10)) & " | " & CStr(info(11)) & " | recorded Addresses=" & addresses.Count
    top = VD_BuildGrid(wb, ws, info)
    top = PutCollection(ws, top, "Fail BLK detail", Array("WT", "Pattern", "Stop TNO", "X", "Y", "Chip Fail BLK Count", "BLK", "FBC", "Pattern Address Count", "Write Fail", "Read Fail", "Write&Read Fail", "No monitor item Fail", "Endpoint Missing Count", "Coverage Basis", "Analysis Status"), blocks, "tblChipBLK")
    addrHeader = top + 1
    top = PutCollection(ws, top, "Fail Bit / Monitor matching", Array("WT", "Pattern", "Stop TNO", "X", "Y", "BLK", "WL", "BL", "IO", "Function Fail TNOs", "Write Monitor TNOs", "Read Monitor TNOs", "Write Monitor Fail", "Read Monitor Fail", "Classification", "Source Row"), addresses, "tblChipMonitorAddr")
    For r = 1 To addresses.Count
        item = addresses(r): ws.Hyperlinks.Add ws.Cells(addrHeader + r, 16), "", "'Fail Address'!A" & CLng(item(15)), , CStr(item(15))
    Next r
    ws.UsedRange.Font.Name = "Calibri": ws.UsedRange.Font.Size = 11: ws.UsedRange.HorizontalAlignment = xlCenter: ws.UsedRange.VerticalAlignment = xlCenter
    ws.Columns("A:P").ColumnWidth = 12: ws.Columns("J:L").ColumnWidth = 24: ws.Columns("O:P").ColumnWidth = 28
    ws.Range("A1:P1").Interior.Color = RGB(23, 50, 77): ws.Range("A1:P1").Font.Color = vbWhite
    If Left$(CStr(ws.Range("A5").Value2), 24) = "BLK Voltage Distribution" Then
        ws.Columns("A:P").ColumnWidth = 12
        ws.Range("A5:P5").RowHeight = 28: ws.Range("A5:P5").Interior.Color = RGB(23, 50, 77): ws.Range("A5:P5").Font.Color = vbWhite
        If Not ws.ListObjects("tblChipMonitorAddr").DataBodyRange Is Nothing Then ws.ListObjects("tblChipMonitorAddr").DataBodyRange.WrapText = True: ws.ListObjects("tblChipMonitorAddr").DataBodyRange.RowHeight = 72
    End If
    ws.Activate: Application.ActiveWindow.DisplayGridlines = False: Application.Goto ws.Range("A1"), True
    BLK_OrderSheets wb
    Exit Sub
Failed:
    MsgBox "Cannot open Chip Fail Details: " & Err.Description, vbExclamation
End Sub

Public Function BLK_SourceRowCell(ByVal ws As Object, ByVal target As Range) As Boolean
    Dim table As ListObject, col As ListColumn
    On Error GoTo Done
    If TypeName(ws) <> "Worksheet" Or target.CountLarge <> 1 Then Exit Function
    If ws.Name <> "Fail Analysis" And ws.Name <> "Chip Fail Details" Then Exit Function
    If Not IsNumeric(target.Value2) Then Exit Function
    If Val(CStr(target.Value2)) < 2 Then Exit Function
    For Each table In ws.ListObjects
        For Each col In table.ListColumns
            If col.Name = "Source Row" Or col.Name = "Source Address Row" Then
                If Not col.DataBodyRange Is Nothing Then If Not Intersect(target, col.DataBodyRange) Is Nothing Then BLK_SourceRowCell = True: Exit Function
            End If
        Next col
    Next table
Done:
End Function

Public Sub BLK_JumpSourceRow(ByVal wb As Workbook, ByVal target As Range)
    Dim oldEvents As Boolean, row As Long
    oldEvents = Application.EnableEvents
    On Error GoTo Done
    row = CLng(target.Value2)
    If row < 2 Or row > wb.Worksheets("Fail Address").Cells(wb.Worksheets("Fail Address").Rows.Count, 1).End(xlUp).Row Then Exit Sub
    Application.EnableEvents = False
    Application.Goto wb.Worksheets("Fail Address").Cells(row, 1), True
Done:
    Application.EnableEvents = oldEvents
End Sub

Public Sub BLK_PrepareEmptyTemplate(ByVal wb As Workbook)
    Dim ws As Worksheet, cfg As Worksheet, locations As Worksheet, index As Worksheet
    Dim tnos As Object, labels As Object, tno As Variant, header() As Variant, i As Long, width As Long, message As String
    If PresentationLastRow(wb.Worksheets("Fail Address"), 1) > 1 Or PresentationLastRow(wb.Worksheets("FBC by BLK"), 2) > 1 Then Err.Raise vbObjectError + 491, , "Template preparation requires an empty workbook."
    Set cfg = wb.Worksheets("Config"): BLK_Setup cfg
    Set tnos = BLK_AddressTNOs(): Set labels = ImportFailAddresses.BuildAddressLabels(cfg)
    Set ws = wb.Worksheets("Fail Address"): width = 7 + tnos.Count
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.UsedRange.Clear
    If ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1 > width Then ws.Range(ws.Columns(width + 1), ws.Columns(ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1)).Delete
    ReDim header(1 To 1, 1 To width)
    header(1, 1) = "X": header(1, 2) = "Y": header(1, 3) = "BLK": header(1, 4) = "WL": header(1, 5) = "BL": header(1, 6) = "IO"
    header(1, 7) = "Cat."
    For Each tno In tnos: header(1, 8 + CLng(tnos(tno))) = "WT1" & vbLf & labels(tno): Next tno
    ws.Cells(1, 1).Resize(1, width).Value2 = header
    With ws.Range("A1:G1"): .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: End With
    With ws.Cells(1, 8).Resize(1, tnos.Count): .Interior.Color = RGB(219, 234, 254): .Font.Bold = True: .Orientation = 90: .WrapText = True: End With
    Set locations = wb.Worksheets("Chip Locations"): locations.Range("A2:F" & locations.UsedRange.Rows.Count + 1).ClearContents
    Set ws = wb.Worksheets("FBC by BLK"): If ws.AutoFilterMode Then ws.AutoFilterMode = False
    If ws.UsedRange.Rows.Count > 1 Then ws.Range("A2:IZ" & ws.UsedRange.Rows.Count).ClearContents
    Set index = wb.Worksheets("Import Index"): index.Range("A2:O" & index.UsedRange.Rows.Count + 1).ClearContents
    message = BLK_BuildComparison(wb)
    If Len(message) > 0 Then Err.Raise vbObjectError + 491, , message
    BLK_ApplyPresentation wb
    cfg.Activate
    BLK_Release
End Sub


Public Function BLK_PhaseCount(ByVal wb As Workbook) As Long
    Dim phase As Long, status As String
    If mAvailablePhases > 0 Then BLK_PhaseCount = mAvailablePhases: Exit Function
    For phase = 1 To 3
        status = CStr(wb.Worksheets("Chip Locations").Cells(phase + 1, 5).Value2)
        If status = "OK" Or status = "Data Error" Then BLK_PhaseCount = phase
    Next phase
    If BLK_PhaseCount = 0 Then BLK_PhaseCount = 1
End Function

Public Sub BLK_InvalidateDetails(ByVal wb As Workbook, ByVal message As String)
    Dim ws As Worksheet, i As Long, owned As Boolean
    On Error Resume Next
    Set ws = wb.Worksheets("Chip Fail Details")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    For i = 1 To ws.CustomProperties.Count
        If ws.CustomProperties(i).Name = "ASCChipDetails" Then owned = True
    Next i
    If Not owned Then Exit Sub
    ws.Range("A3").Value2 = message
    ws.Range("A3:M3").Interior.Color = RGB(255, 217, 102)
End Sub

Public Sub BLK_OrderSheets(ByVal wb As Workbook)
    Dim info As Worksheet
    On Error Resume Next
    Set info = wb.Worksheets("Import Info")
    On Error GoTo 0
    If Not info Is Nothing Then
        If info.Index <> wb.Worksheets.Count Then info.Move After:=wb.Worksheets(wb.Worksheets.Count)
    Else
        If wb.Worksheets("Log").Visible = xlSheetVisible And wb.Worksheets("Log").Index <> wb.Worksheets.Count Then wb.Worksheets("Log").Move After:=wb.Worksheets(wb.Worksheets.Count)
    End If
End Sub

Public Sub BLK_ApplyPresentation(ByVal wb As Workbook)
    Dim previousWindow As Window, previousSheet As Object, window As Window, ws As Worksheet, lastRow As Long, lastCol As Long
    On Error GoTo Restore
    Set previousWindow = Application.ActiveWindow: Set previousSheet = Application.ActiveSheet
    wb.Worksheets("Config").Range("H1:N901").Borders.LineStyle = xlNone
    If wb.Worksheets("Log").Visible = xlSheetVisible And wb.Worksheets("Log").Index <> wb.Worksheets.Count Then wb.Worksheets("Log").Move After:=wb.Worksheets(wb.Worksheets.Count)
    If wb.Worksheets("Chip Locations").Visible = xlSheetVisible And wb.Worksheets("Chip Locations").Index <> wb.Worksheets.Count - 1 Then wb.Worksheets("Chip Locations").Move Before:=wb.Worksheets("Log")
    Set ws = wb.Worksheets("FBC by BLK")
    lastRow = PresentationLastRow(ws, 2)
    If lastRow > 1 Then ws.Cells(2, 5).Resize(lastRow - 1, 256).NumberFormat = "0;-0;;@"
    BLK_SimplifyButtons wb
    For Each window In wb.Windows
        window.Activate
        For Each ws In wb.Worksheets
            If ws.Visible = xlSheetVisible Then
                BLK_SizeOutput ws
                ws.Activate: window.DisplayGridlines = False
                lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
                lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
                If ws.Name = "FBC by BLK" Then lastRow = PresentationLastRow(ws, 2): lastCol = 260
                If ws.Name = "Fail Address" Then
                    lastRow = PresentationLastRow(ws, 1)
                    BLK_FormatFailAddress wb
                End If
                If lastRow < 1 Then lastRow = 1
                If ws.Name <> "Config" Then
                    ws.Cells(1, 1).Resize(1, lastCol).Borders.LineStyle = xlContinuous
                    ws.Cells(1, 1).Resize(lastRow, lastCol).BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(80, 80, 80)
                End If
            End If
        Next ws
    Next window
    wb.Worksheets("Config").Range("H1:N901").Borders.LineStyle = xlNone
    If wb.Worksheets("Log").Visible = xlSheetVisible And wb.Worksheets("Log").Index <> wb.Worksheets.Count Then wb.Worksheets("Log").Move After:=wb.Worksheets(wb.Worksheets.Count)
    If wb.Worksheets("Chip Locations").Visible = xlSheetVisible And wb.Worksheets("Chip Locations").Index <> wb.Worksheets.Count - 1 Then wb.Worksheets("Chip Locations").Move Before:=wb.Worksheets("Log")
Restore:
    On Error Resume Next
    If Not previousWindow Is Nothing Then previousWindow.Activate
    If Not previousSheet Is Nothing Then previousSheet.Activate
End Sub

Public Sub BLK_SizeOutput(ByVal ws As Worksheet)
    Dim lastRow As Long, lastCol As Long, c As Long, firstData As Long, height As Double, maximum As Double
    If ws.Name = "Chip Fail Details" Then
        If VD_Property(ws, "ASCVoltageGrid") = "1" And Left$(CStr(ws.Range("A5").Value2), 24) = "BLK Voltage Distribution" Then Exit Sub
    End If
    If ws.Name = "Config" Or ws.Name = "Import Index" Or ws.Name = "FBC by BLK" Or ws.Name = "Fail Analysis" Or ws.Name = "Import Info" Or ws.Name = "Wafer Maps" Then Exit Sub
    If Right$(ws.Name, 8) = "Fail Map" Or ws.Name = "FailBLK Loss Map" Or ws.Name = "W&R Map" Then Exit Sub
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If ws.Name = "BLK_FBC_Comparison" And lastCol > 3 Then ws.Cells(2, 4).Resize(1, lastCol - 3).Orientation = 90
    If ws.Name = "Fail Address" And lastCol > 7 Then ws.Cells(1, 8).Resize(1, lastCol - 7).Orientation = 90
    With ws.UsedRange
        .Font.Name = "Calibri": .Font.Size = 11: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Columns.AutoFit
    End With
    For c = 1 To lastCol
        maximum = 32
        If ws.Name = "Log" And c = 4 Then maximum = 100
        If ws.Name = "Chip Locations" And c = 6 Then maximum = 70
        If ws.Name = "Missing BLK Address" Then maximum = 45
        If ws.Name = "BLK_FBC_Comparison" Then maximum = 26
        If ws.Columns(c).ColumnWidth < 8 Then ws.Columns(c).ColumnWidth = 8
        If ws.Columns(c).ColumnWidth > maximum Then ws.Columns(c).ColumnWidth = maximum
    Next c
    firstData = 2
    Select Case ws.Name
        Case "BLK_FBC_Comparison"
            firstData = 3
            For c = 4 To lastCol
                If ws.Cells(2, c).Orientation = 90 Or ws.Cells(2, c).Orientation = xlUpward Then
                    If ws.Columns(c).ColumnWidth < 7 Then ws.Columns(c).ColumnWidth = 7
                    If ws.Columns(c).ColumnWidth > 14 Then ws.Columns(c).ColumnWidth = 14
                End If
            Next c
            ws.Rows(1).RowHeight = 30
            ws.Rows(2).AutoFit
            height = ws.Rows(2).RowHeight
            If height < 90 Then height = 90
            If height > 200 Then height = 200
            ws.Rows(2).RowHeight = height
        Case "Fail Address"
            For c = 7 To lastCol
                If ws.Columns(c).ColumnWidth > 12 Then ws.Columns(c).ColumnWidth = 12
            Next c
            ws.Rows(1).AutoFit
            height = ws.Rows(1).RowHeight
            If height < 90 Then height = 90
            If height > 200 Then height = 200
            ws.Rows(1).RowHeight = height
        Case "Monitor Fail Analysis", "Monitor BLK Summary"
            firstData = 5
            ws.Rows(1).RowHeight = 30: ws.Rows(2).RowHeight = 42: ws.Rows(3).RowHeight = 30
            ws.Rows(4).RowHeight = 160
        Case "Missing BLK Address"
            firstData = 5
            ws.Rows(1).RowHeight = 30: ws.Rows(2).RowHeight = 42: ws.Rows(3).RowHeight = 30
            ws.Rows(4).AutoFit
            If ws.Rows(4).RowHeight < 30 Then ws.Rows(4).RowHeight = 30
        Case "Log"
            ws.Range("D2:D5").WrapText = True
            ws.Rows("2:5").AutoFit
            For c = 2 To 5
                If ws.Rows(c).RowHeight < 26 Then ws.Rows(c).RowHeight = 26
                If ws.Rows(c).RowHeight > 180 Then ws.Rows(c).RowHeight = 180
            Next c
            ws.Rows(1).RowHeight = 30
            Exit Sub
        Case Else
            ws.Rows(1).AutoFit
            If ws.Rows(1).RowHeight < 30 Then ws.Rows(1).RowHeight = 30
    End Select
    If lastRow >= firstData Then ws.Cells(firstData, 1).Resize(lastRow - firstData + 1, lastCol).RowHeight = 26
    If ws.Name = "Monitor Fail Analysis" And lastRow >= 5 Then
        ws.Range("J5:L" & lastRow).WrapText = True
        ws.Cells(5, 1).Resize(lastRow - 4, lastCol).RowHeight = 60
    End If
    If ws.Name = "Chip Locations" Then
        ws.Range("F2:F4").WrapText = True: ws.Rows("2:4").AutoFit
        For c = 2 To 4
            If ws.Rows(c).RowHeight < 26 Then ws.Rows(c).RowHeight = 26
            If ws.Rows(c).RowHeight > 120 Then ws.Rows(c).RowHeight = 120
        Next c
    End If
End Sub

Public Sub BLK_SimplifyButtons(ByVal wb As Workbook)
    Dim buttons As Object, button As Object, i As Long, action As String, found As Boolean
    Set buttons = wb.Worksheets("Config").Buttons
    For i = 1 To buttons.Count
        If InStr(1, CStr(buttons.Item(i).OnAction), "ImportASCAddressOnly", vbTextCompare) > 0 Then found = True
    Next i
    If Not found Then
        With wb.Worksheets("Config")
            Set button = buttons.Add(.Range("H14").Left, .Range("H14").Top, .Range("H14:K15").Width, .Range("H14:K15").Height)
        End With
        button.Name = "btnImportAddressOnly": button.OnAction = "ImportASCAddressOnly"
        If buttons.Count > 0 Then
            button.Left = buttons.Item(1).Left: button.Width = buttons.Item(1).Width: button.Height = buttons.Item(1).Height
        End If
        Set buttons = wb.Worksheets("Config").Buttons
    End If
    For i = 1 To buttons.Count
        Set button = buttons.Item(i): action = CStr(button.OnAction)
        If InStr(1, action, "ImportASCFailAddresses", vbTextCompare) > 0 Then button.Caption = ChrW(&H532F) & ChrW(&H5165) & ChrW(&H4E26) & ChrW(&H5206) & ChrW(&H6790)
        If InStr(1, action, "ImportASCFBC", vbTextCompare) > 0 Then button.Caption = ChrW(&H532F) & ChrW(&H5165) & " FBC"
        If InStr(1, action, "BuildFailMaps", vbTextCompare) > 0 Then button.Caption = ChrW(&H66F4) & ChrW(&H65B0) & " Map"
        If InStr(1, action, "FindMissingBLKAddresses", vbTextCompare) > 0 Then button.Caption = ChrW(&H6838) & ChrW(&H5C0D) & " BLK Address"
        If InStr(1, action, "ImportASCAddressOnly", vbTextCompare) > 0 Then button.Caption = ChrW(&H50C5) & ChrW(&H532F) & ChrW(&H5165) & " Address"
        button.Font.Name = "Calibri": button.Font.Size = 11
        button.HorizontalAlignment = xlCenter: button.VerticalAlignment = xlCenter
    Next i
End Sub


Private Function PresentationLastRow(ByVal ws As Worksheet, ByVal column As Long) As Long
    Dim last As Long, first As Long, take As Long, i As Long, data As Variant, value As Variant
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    Do While last >= 1
        first = last - 8191: If first < 1 Then first = 1
        take = last - first + 1: data = ws.Cells(first, column).Resize(take, 1).Value2
        For i = take To 1 Step -1
            If take = 1 Then value = data Else value = data(i, 1)
            If IsError(value) Then PresentationLastRow = first + i - 1: Exit Function
            If Len(CStr(value)) > 0 Then PresentationLastRow = first + i - 1: Exit Function
        Next i
        last = first - 1
    Loop
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

Public Sub BLK_FormatFailAddress(ByVal wb As Workbook)
    Dim ws As Worksheet, previousSheet As Object, previousWindow As Window, window As Window
    Dim lastRow As Long, lastCol As Long
    Set ws = wb.Worksheets("Fail Address")
    lastRow = PresentationLastRow(ws, 1): If lastRow < 1 Then lastRow = 1
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column: If lastCol < 7 Then lastCol = 7
    With ws.Cells(1, 1).Resize(lastRow, lastCol).Borders
        .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(191, 191, 191)
    End With
    With ws.Cells(1, 7).Resize(lastRow, 1).Borders(xlEdgeRight)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(80, 80, 80)
    End With
    Set previousWindow = Application.ActiveWindow: Set previousSheet = Application.ActiveSheet
    For Each window In wb.Windows
        window.Activate: ws.Activate
        window.FreezePanes = False: window.SplitColumn = 7: window.SplitRow = 1
        window.FreezePanes = True: window.DisplayGridlines = False
        window.ScrollRow = 1: window.ScrollColumn = 1
    Next window
    If Not previousWindow Is Nothing Then previousWindow.Activate
    If Not previousSheet Is Nothing Then previousSheet.Activate
End Sub

Private Sub RevealMapSources(ByVal wb As Workbook, ByVal show As Boolean)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name = "W&R Map" Or ws.Name = "FailBLK Loss Map" Or Right$(ws.Name, 8) = "Fail Map" Then
            If show Then ws.Visible = xlSheetVisible Else ws.Visible = xlSheetVeryHidden
        End If
    Next ws
End Sub

Private Sub MapNavButton(ByVal ws As Worksheet, ByVal name As String, ByVal caption As String, ByVal left As Double, ByVal top As Double, ByVal width As Double, ByVal address As String, Optional ByVal action As String = "GoWaferMap")
    Dim button As Object
    Set button = ws.Buttons.Add(left, top, width, 24)
    button.Name = name: button.Caption = caption: button.OnAction = action
    button.ShapeRange.AlternativeText = address
    button.Font.Name = "Calibri": button.Font.Size = 11
    button.HorizontalAlignment = xlCenter: button.VerticalAlignment = xlCenter
End Sub

Public Sub BLK_GoWaferMap(ByVal wb As Workbook, Optional ByVal buttonName As String = "")
    Dim ws As Worksheet, address As String
    Set ws = wb.Worksheets("Wafer Maps")
    If Len(buttonName) = 0 Then buttonName = CStr(Application.Caller)
    address = ws.Shapes(buttonName).AlternativeText
    If Len(address) = 0 Then Exit Sub
    ws.Activate: Application.Goto ws.Range(address), True
    Application.ActiveWindow.ScrollColumn = 1
End Sub

Public Sub BLK_BuildUnifiedMaps(ByVal wb As Workbook)
    Dim ws As Worksheet, loc As Worksheet, idx As Worksheet, source As Worksheet
    Dim coords As Object, patterns As Object, wanted As Object, hard As Object, pass As Object, counts As Object
    Dim ordered As Collection, item As Variant, xy As Variant, data As Variant, mapIndex As Variant
    Dim kind As Variant, name As Variant, detail As Variant, keys As Variant, key As String, label As String
    Dim last As Long, i As Long, j As Long, phase As Long, phases As Long, slot As Long, panelDest As Long, mapDest As Long
    Dim minX As Long, maxX As Long, minY As Long, maxY As Long, x As Long, y As Long, r As Long, c As Long
    Dim span As Long, height As Long, top As Long, sourceTop As Long, sourceLeft As Long, patternIndex As Long, n As Long
    Dim classNames As Variant, v As Long, keep As Long, packet() As Variant
    Set idx = wb.Worksheets("Import Index"): Set loc = wb.Worksheets("Chip Locations")
    Set ws = ReportSheet(wb, "Wafer Maps")
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 6) = "navMap" Then ws.Shapes(i).Delete
    Next i
    ws.Columns("A:V").ColumnWidth = 6.5: ws.Rows("1:8").RowHeight = 26
    ws.Range("A1:U1").Merge: ws.Range("A1").Value2 = "Wafer Maps"
    With ws.Range("A1:U1"): .Font.Name = "Calibri": .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Size = 24: .Font.Bold = True: .RowHeight = 36: .HorizontalAlignment = xlCenter: End With
    ws.Range("A2:U2").Merge: ws.Range("A2").Value2 = "Fail BLK / W&R matching / Fail BLK loss"
    MapNavButton ws, "navMapPPT", "Export PPT", ws.Range("A7").Left, ws.Range("A7").Top, 145, "", "ExportWaferMapsPPT"
    MapNavButton ws, "navMapRefresh", "Refresh Maps", ws.Range("E7").Left, ws.Range("E7").Top, 145, "", "BuildFailMaps"
    ws.Range("I7:U7").Merge: ws.Range("I7").Value2 = "Double-click Chip for details"
    last = idx.Cells(idx.Rows.Count, 46).End(xlUp).Row: idx.Cells(1, 46).Resize(last, 8).ClearContents
    idx.Cells(1, 46).Resize(1, 8).Value2 = Array("Sheet", "Kind", "WT", "Pattern", "Top", "Left", "Rows", "Columns"): panelDest = 2
    last = loc.Cells(loc.Rows.Count, 1).End(xlUp).Row
    If last < 2 Then GoTo Finish
    If InStr(1, CStr(wb.Worksheets("Monitor Fail Analysis").Range("A2").Value2), "Source updated", vbTextCompare) > 0 Then Err.Raise vbObjectError + 520, , "Refresh Monitor analysis before merging Maps."
    Set coords = CreateObject("Scripting.Dictionary"): data = loc.Range("A2:B" & last).Value2
    For i = 1 To UBound(data, 1): coords(CStr(data(i, 1)) & "|" & CStr(data(i, 2))) = True: Next i
    Set ordered = SortedDUTs(coords): xy = Split(CStr(ordered(1)), "|")
    minX = CLng(xy(0)): maxX = minX: minY = CLng(xy(1)): maxY = minY
    For Each item In ordered
        xy = Split(CStr(item), "|"): x = CLng(xy(0)): y = CLng(xy(1))
        If x < minX Then minX = x
        If x > maxX Then maxX = x
        If y < minY Then minY = y
        If y > maxY Then maxY = y
    Next item
    span = 2 * (maxX - minX + 1) + 2: height = 2 * (maxY - minY + 1) + 2
    phases = BLK_PhaseCount(wb): Set patterns = ReadStopPatterns(wb.Worksheets("Config"), wanted, hard)
    If patterns.Count = 0 Then GoTo Finish
    keys = patterns.Keys: Set pass = ConfirmedFunctionPassChips(wb, coords, phases)
    Set counts = CreateObject("Scripting.Dictionary"): Set source = wb.Worksheets("Monitor Fail Analysis")
    last = source.Cells(source.Rows.Count, 1).End(xlUp).Row
    If last >= 5 Then
        data = source.Range("A5:P" & last).Value2
        For i = 1 To UBound(data, 1)
            key = CStr(data(i, 1)) & "|" & CStr(data(i, 2)) & "|" & CStr(data(i, 4)) & "|" & CStr(data(i, 5)) & "|" & CStr(data(i, 15))
            If Not counts.Exists(key) Then counts(key) = 0
            counts(key) = CLng(counts(key)) + 1
        Next i
    End If
    last = idx.Cells(idx.Rows.Count, 33).End(xlUp).Row
    If last < 2 Then Err.Raise vbObjectError + 520, , "Refresh Maps before merging them."
    mapIndex = idx.Cells(2, 33).Resize(last - 1, 12).Value2
    ReDim packet(1 To UBound(mapIndex, 1), 1 To 12)
    For i = 1 To UBound(mapIndex, 1)
        If CStr(mapIndex(i, 1)) <> "Wafer Maps" Then
            keep = keep + 1: For j = 1 To 12: packet(keep, j) = mapIndex(i, j): Next j
        End If
    Next i
    idx.Cells(2, 33).Resize(last - 1, 12).ClearContents
    If keep > 0 Then idx.Cells(2, 33).Resize(UBound(packet, 1), 12).Value2 = packet
    mapDest = keep + 2: top = 9
    classNames = Array("Read Fail", "Write Fail", "Write&Read Fail", "No monitor item Fail")
    For Each kind In Array("COUNT", "QUAD", "LOSS")
        slot = 0
        For phase = 1 To phases
            For patternIndex = 0 To patterns.Count - 1
                name = keys(patternIndex): detail = patterns(name): slot = slot + 1
                If top + height > ws.Rows.Count Then Err.Raise vbObjectError + 520, , "Merged Maps exceed Excel row limit."
                sourceLeft = 1 + (phase - 1) * span: sourceTop = 3
                Select Case CStr(kind)
                    Case "COUNT": Set source = wb.Worksheets(FailMapNameForPattern(wb, CStr(name))): label = "Fail BLK"
                    Case "LOSS": Set source = wb.Worksheets("FailBLK Loss Map"): sourceTop = 3 + patternIndex * (height + 4): label = "Fail BLK / Loss BLK"
                    Case "QUAD": Set source = wb.Worksheets("W&R Map"): label = "W&R"
                End Select
                source.Cells(sourceTop, sourceLeft).Resize(height, span - 1).Copy Destination:=ws.Cells(top, 1)
                With ws.Cells(top, 2).MergeArea.Cells(1, 1)
                    .Value2 = "WT" & phase & " " & CStr(name) & " - " & label: .Font.Size = 24
                End With
                ws.Rows(top).Resize(height).RowHeight = 36
                If CStr(kind) = "QUAD" Then
                    For Each item In ordered
                        xy = Split(CStr(item), "|"): x = CLng(xy(0)): y = CLng(xy(1))
                        r = top + 2 + 2 * (y - minY): c = 2 + 2 * (x - minX)
                        With ws.Cells(r, c).Resize(2, 2)
                            .UnMerge: .ClearContents: .Interior.Pattern = xlNone: .Font.Color = vbBlack
                        End With
                        key = "WT" & phase & "|" & CStr(item)
                        If pass.Exists(key) Then
                            With ws.Cells(r, c).Resize(2, 2): .Merge: .Interior.Color = RGB(0, 176, 240): End With
                            IndexMapCell idx, mapDest, ws, r, c, "WT" & phase, CStr(name), x, y, "Confirmed Pass", "", CStr(detail(0)), 0, "All Function Pass"
                        Else
                            For j = 0 To 3
                                key = "WT" & phase & "|" & CStr(name) & "|" & CStr(item) & "|" & CStr(classNames(j)): v = 0
                                If counts.Exists(key) Then v = CLng(counts(key))
                                With ws.Cells(r + j \ 2, c + j Mod 2)
                                    If v > 0 Then
                                        .Value2 = v
                                        If j = 1 Then .Interior.Color = RGB(255, 217, 202)
                                        If j = 2 Then .Interior.Color = RGB(255, 102, 0)
                                        If j = 3 Then .Font.Color = vbWhite
                                    End If
                                End With
                                IndexMapCell idx, mapDest, ws, r + j \ 2, c + j Mod 2, "WT" & phase, CStr(name), x, y, "Monitor Address Count", CStr(classNames(j)), CStr(detail(0)), v, IIf(v > 0, "Recorded matches", "No recorded matches")
                            Next j
                        End If
                        ChipBorder ws.Cells(r, c).Resize(2, 2)
                    Next item
                Else
                    For i = 1 To UBound(mapIndex, 1)
                        If CStr(mapIndex(i, 1)) = source.Name And CStr(mapIndex(i, 4)) = "WT" & phase And CStr(mapIndex(i, 5)) = CStr(name) Then
                            r = top + CLng(mapIndex(i, 2)) - sourceTop: c = 1 + CLng(mapIndex(i, 3)) - sourceLeft
                            If r >= top And r < top + height And c >= 1 And c < span Then
                                IndexMapCell idx, mapDest, ws, r, c, CStr(mapIndex(i, 4)), CStr(name), CLng(mapIndex(i, 6)), CLng(mapIndex(i, 7)), CStr(mapIndex(i, 8)), CStr(mapIndex(i, 9)), CStr(mapIndex(i, 10)), mapIndex(i, 11), CStr(mapIndex(i, 12))
                            End If
                        End If
                    Next i
                End If
                idx.Cells(panelDest, 46).Resize(1, 8).Value2 = Array("Wafer Maps", kind, "WT" & phase, name, top, 1, height, span - 1): panelDest = panelDest + 1
                n = 3: If CStr(kind) = "QUAD" Then n = 4
                If CStr(kind) = "LOSS" Then n = 5
                MapNavButton ws, "navMap_" & kind & "_" & slot, "WT" & phase & " " & CStr(name) & " " & label, ws.Cells(n, 1).Left + (slot - 1) * 130, ws.Cells(n, 1).Top, 126, "A" & top
                top = top + height + 4
            Next patternIndex
        Next phase
    Next kind
Finish:
    With ws.Range("A2:U8"): .Font.Name = "Calibri": .Font.Size = 11: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: End With
    ws.Activate
    With Application.ActiveWindow
        .DisplayGridlines = False: .FreezePanes = False: .SplitRow = 8: .SplitColumn = 0: .FreezePanes = True
        .ScrollRow = 1: .ScrollColumn = 1: .Zoom = 70
    End With
    RevealMapSources wb, False
    BLK_OrderSheets wb
End Sub

Private Function PPTBitmap(ByVal slide As Object, ByVal source As Range, ByVal left As Double, ByVal top As Double, ByVal width As Double, ByVal name As String) As Object
    Dim pasted As Object, pic As Object, attempt As Long, errorText As String
    For attempt = 1 To 4
        On Error Resume Next
        Err.Clear: source.Worksheet.Activate
        source.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
        DoEvents
        If Err.Number = 0 Then Set pasted = slide.Shapes.PasteSpecial(1)
        errorText = Err.Description
        On Error GoTo 0
        If Not pasted Is Nothing Then Exit For
        Application.Wait Now + TimeSerial(0, 0, 1)
    Next attempt
    If pasted Is Nothing Then Err.Raise vbObjectError + 521, , "Cannot copy Map bitmap: " & errorText
    Set pic = pasted.Item(1): pic.Name = name: pic.LockAspectRatio = -1
    pic.Width = width: pic.Left = left: pic.Top = top
    Set PPTBitmap = pic
End Function

Private Sub PPTLabel(ByVal slide As Object, ByVal text As String, ByVal left As Double, ByVal top As Double, ByVal width As Double, ByVal height As Double, ByVal size As Double)
    Dim box As Object
    Set box = slide.Shapes.AddTextbox(1, left, top, width, height)
    With box.TextFrame
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        With .TextRange: .Text = text: .Font.Name = "Calibri": .Font.Size = size: .Font.Color.RGB = RGB(23, 50, 77): End With
    End With
End Sub

Public Function BLK_ExportMapsPPT(ByVal wb As Workbook) As String
    Const PT_PER_CM As Double = 28.3464566929134
    Dim pp As Object, deck As Object, slide As Object, pic As Object, ws As Worksheet, idx As Worksheet, previous As Object
    Dim data As Variant, panels As Object, order As New Collection, entry As Variant, item As Variant, kind As Variant
    Dim i As Long, last As Long, columns As Long, pictureLeft As Double, width As Double, top As Double, mainHeight As Double
    Dim slideWidth As Double, slideHeight As Double, lowerTop As Double, insetTop As Double, insetWidth As Double, h As Double
    Dim key As String, base As String, path As String, suffix As Long, oldScreen As Boolean, oldEvents As Boolean, created As Boolean
    Dim errorNumber As Long, errorText As String
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    On Error GoTo Failed
    Set ws = wb.Worksheets("Wafer Maps"): Set idx = wb.Worksheets("Import Index")
    If InStr(1, CStr(ws.Range("A2").Value2), "updated", vbTextCompare) > 0 Then Err.Raise vbObjectError + 522, , "Refresh Maps after source changes before exporting PPT."
    last = idx.Cells(idx.Rows.Count, 46).End(xlUp).Row
    If last < 2 Then Err.Raise vbObjectError + 522, , "Import Raw Data and build Maps before exporting PPT."
    data = idx.Range("AT2:BA" & last).Value2: Set panels = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(data, 1)
        key = CStr(data(i, 2)) & "|" & CStr(data(i, 3)) & "|" & CStr(data(i, 4))
        panels(key) = Array(CLng(data(i, 5)), CLng(data(i, 6)), CLng(data(i, 7)), CLng(data(i, 8)))
        If CStr(data(i, 2)) = "COUNT" Then order.Add CStr(data(i, 3)) & "|" & CStr(data(i, 4))
    Next i
    columns = order.Count: If columns = 0 Then Err.Raise vbObjectError + 522, , "No Map panels to export."
    Set previous = Application.ActiveSheet
    Application.ScreenUpdating = True: Application.EnableEvents = False
    On Error Resume Next
    Set pp = GetObject(, "PowerPoint.Application")
    On Error GoTo Failed
    If pp Is Nothing Then Set pp = CreateObject("PowerPoint.Application"): created = True
    pp.Visible = True
    Set deck = pp.Presentations.Add
    slideWidth = 37 * PT_PER_CM: width = (slideWidth - 36 - 6 * (columns - 1)) / columns
    entry = panels("COUNT|" & CStr(order(1)))
    With ws.Cells(CLng(entry(0)), CLng(entry(1))).Resize(CLng(entry(2)), CLng(entry(3)))
        mainHeight = width * .Height / .Width
    End With
    h = (slideWidth * 9 / 16 - 136) / 2.36
    If mainHeight > h Then width = width * h / mainHeight: mainHeight = h
    lowerTop = 54 + mainHeight + 34: insetTop = lowerTop + mainHeight + 6: insetWidth = width * 0.36
    slideHeight = insetTop + mainHeight * 0.36 + 48
    If slideHeight < slideWidth * 9 / 16 Then slideHeight = slideWidth * 9 / 16
    deck.PageSetup.SlideWidth = slideWidth: deck.PageSetup.SlideHeight = slideHeight
    Set slide = deck.Slides.Add(1, 12)
    PPTLabel slide, "Fail BLK Count", 18, 12, slideWidth - 36, 30, 24
    PPTLabel slide, "W&R Monitor Matching", 18, lowerTop - 29, slideWidth - 36, 26, 20
    For i = 1 To columns
        item = order(i): pictureLeft = (slideWidth - (columns * width + 6 * (columns - 1))) / 2 + (i - 1) * (width + 6)
        For Each kind In Array("COUNT", "QUAD", "LOSS")
            key = CStr(kind) & "|" & CStr(item)
            If Not panels.Exists(key) Then Err.Raise vbObjectError + 522, , "Missing PPT panel: " & key
            entry = panels(key): top = 54: h = width
            If CStr(kind) = "QUAD" Then top = lowerTop
            If CStr(kind) = "LOSS" Then top = insetTop: h = insetWidth
            Set pic = PPTBitmap(slide, ws.Cells(CLng(entry(0)), CLng(entry(1))).Resize(CLng(entry(2)), CLng(entry(3))), pictureLeft + IIf(CStr(kind) = "LOSS", width - insetWidth, 0), top, h, CStr(kind) & "_" & Replace(CStr(item), "|", "_"))
            pic.AlternativeText = "WT / Pattern: " & CStr(item) & "; Map: " & CStr(kind) & "; Excel: " & ws.Cells(CLng(entry(0)), CLng(entry(1))).Address(False, False)
        Next kind
    Next i
    PPTLabel slide, "W&R: Read only / Write only / Both / No recorded Monitor match. Small inset: Fail BLK / Loss BLK.", 18, slideHeight - 30, slideWidth - 36, 24, 11
    base = wb.Name: If InStrRev(base, ".") > 0 Then base = Left$(base, InStrRev(base, ".") - 1)
    path = wb.Path: If Len(path) = 0 Then path = Application.DefaultFilePath
    base = path & Application.PathSeparator & base & "_Wafer_Maps": path = base & ".pptx"
    Do While Len(Dir$(path)) > 0
        suffix = suffix + 1: path = base & "_" & Format$(suffix, "000") & ".pptx"
    Loop
    deck.SaveAs path, 24
    BLK_ExportMapsPPT = path
Cleanup:
    On Error Resume Next
    Application.CutCopyMode = False
    If Not previous Is Nothing Then previous.Activate
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    If errorNumber <> 0 Then
        If Not deck Is Nothing Then deck.Close
        If created And Not pp Is Nothing Then pp.Quit
        On Error GoTo 0
        Err.Raise errorNumber, , errorText
    End If
    Exit Function
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Resume Cleanup
End Function

Private Function VD_Voltage(ByVal value As Variant) As Variant
    Dim text As String
    If IsError(value) Or IsEmpty(value) Then Exit Function
    text = Trim$(CStr(value)): If Len(text) = 0 Then Exit Function
    text = Replace(LCase$(text), "v", "")
    If Not IsNumeric(text) Then Exit Function
    If CDbl(text) < 0 Then Exit Function
    VD_Voltage = CDbl(text)
End Function

Private Function VD_VoltageText(ByVal value As Variant) As String
    If IsEmpty(value) Then VD_VoltageText = "V?": Exit Function
    If Abs(CDbl(value) - Round(CDbl(value), 2)) > 0.0000001 Then
        VD_VoltageText = Format$(CDbl(value), "0.000") & "v"
    Else
        VD_VoltageText = Format$(CDbl(value), "0.00") & "v"
    End If
End Function

Private Function VD_Property(ByVal ws As Worksheet, ByVal name As String) As String
    Dim i As Long
    For i = 1 To ws.CustomProperties.Count
        If ws.CustomProperties(i).Name = name Then VD_Property = CStr(ws.CustomProperties(i).Value): Exit Function
    Next i
End Function

Private Function VD_Scope(ByVal cfg As Worksheet, ByVal pattern As String) As Collection
    Dim result As New Collection, data As Variant, i As Long, previous As Long, ending As Long, tno As String
    data = cfg.Range("D2:G901").Value2
    For i = 1 To UBound(data, 1)
        If StrComp(Trim$(CStr(data(i, 4))), pattern, vbTextCompare) = 0 Then ending = i: Exit For
        If Len(Trim$(CStr(data(i, 4)))) > 0 Then previous = i
    Next i
    If ending = 0 Then Err.Raise vbObjectError + 530, , "Pattern configuration changed. Refresh Maps."
    For i = previous + 1 To ending
        tno = Trim$(CStr(data(i, 1)))
        If Len(tno) > 0 Then result.Add Array(CStr(CLng(tno)), VD_Voltage(data(i, 3)), CStr(data(i, 2)))
    Next i
    Set VD_Scope = result
End Function

Private Function VD_State(ByVal counts As Object, ByVal wt As String, ByVal dut As String, ByVal endpoint As String, ByVal blk As Long, ByVal hard As Object, ByVal wanted As Object, ByVal predecessors As Object) As String
    Dim phase As Long, current As Long, tno As Variant, key As String, mask As String, earlier As Collection
    key = wt & "|" & endpoint & "|" & dut
    If Not counts.Exists(key) Then VD_State = "No Data": Exit Function
    current = CLng(Right$(wt, 1))
    For phase = 1 To current
        For Each tno In hard
            key = "WT" & phase & "|" & CStr(tno) & "|" & dut
            If Not counts.Exists(key) Then VD_State = "No Data": Exit Function
            If Mid$(CStr(counts(key)), blk + 1, 1) = "1" Then VD_State = "Excluded: Hard Fail WT" & phase: Exit Function
        Next tno
        If phase < current Then
            For Each tno In wanted
                key = "WT" & phase & "|" & CStr(tno) & "|" & dut
                If Not counts.Exists(key) Then VD_State = "No Data": Exit Function
                If Mid$(CStr(counts(key)), blk + 1, 1) = "1" Then VD_State = "Excluded: WT" & phase & " TNO" & tno: Exit Function
            Next tno
        End If
    Next phase
    If predecessors.Exists(endpoint) Then
        Set earlier = predecessors(endpoint)
        For Each tno In earlier
            key = wt & "|" & CStr(tno) & "|" & dut
            If Not counts.Exists(key) Then VD_State = "No Data": Exit Function
            If Mid$(CStr(counts(key)), blk + 1, 1) = "1" Then VD_State = "Excluded: earlier Pattern TNO" & tno: Exit Function
        Next tno
    End If
    key = wt & "|" & endpoint & "|" & dut
    If Mid$(CStr(counts(key)), blk + 1, 1) = "1" Then VD_State = "Fail" Else VD_State = "Pass"
End Function

Private Function VD_MaxObserved(ByVal list As String, ByVal voltages As Object) As Variant
    Dim token As Variant, key As String, v As Variant, found As Boolean, maximum As Double
    If Len(Trim$(list)) = 0 Then Exit Function
    For Each token In Split(list, ",")
        key = Trim$(CStr(token))
        If Not voltages.Exists(key) Then Exit Function
        v = voltages(key): If IsEmpty(v) Then Exit Function
        If Not found Or CDbl(v) > maximum Then maximum = CDbl(v)
        found = True
    Next token
    If found Then VD_MaxObserved = maximum
End Function

Private Function VD_AddressClass(ByVal writeTnos As String, ByVal readTnos As String, ByVal voltages As Object, ByRef w As Variant, ByRef r As Variant) As String
    w = VD_MaxObserved(writeTnos, voltages): r = VD_MaxObserved(readTnos, voltages)
    If IsEmpty(w) Or IsEmpty(r) Then VD_AddressClass = "Insufficient": Exit Function
    If Abs(CDbl(w) - CDbl(r)) < 0.0000001 Then
        VD_AddressClass = "Equal"
    ElseIf CDbl(w) > CDbl(r) Then
        VD_AddressClass = "Write worse"
    Else
        VD_AddressClass = "Read worse"
    End If
End Function

Private Function VD_MonitorVoltages(ByVal wb As Workbook) As Object
    Dim result As Object, data As Variant, i As Long, tno As String
    Set result = CreateObject("Scripting.Dictionary"): data = wb.Worksheets("Config").Range("A2:C901").Value2
    For i = 1 To UBound(data, 1)
        tno = Trim$(CStr(data(i, 1)))
        If Len(tno) > 0 Then result(CStr(CLng(tno))) = VD_Voltage(data(i, 3))
    Next i
    Set VD_MonitorVoltages = result
End Function

Private Function VD_Stats(ByVal wb As Workbook, ByVal wt As String, ByVal x As Long, ByVal y As Long, ByVal pattern As String) As Object
    Dim result As Object, seen As Object, volts As Object, ws As Worksheet, data As Variant, a As Variant
    Dim i As Long, last As Long, key As String, address As String, category As String, w As Variant, r As Variant, n As Long
    Dim chunkStart As Long, chunkSize As Long
    Set result = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary"): Set volts = VD_MonitorVoltages(wb)
    Set ws = wb.Worksheets("Monitor Fail Analysis"): last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last >= 5 Then
        For chunkStart = 5 To last Step 512
        chunkSize = last - chunkStart + 1: If chunkSize > 512 Then chunkSize = 512
        data = ws.Cells(chunkStart, 1).Resize(chunkSize, 16).Value2
        For i = 1 To UBound(data, 1)
            If CStr(data(i, 1)) = wt And CStr(data(i, 2)) = pattern And CLng(data(i, 4)) = x And CLng(data(i, 5)) = y Then
                key = CStr(data(i, 6)): address = key & "|" & data(i, 7) & "|" & data(i, 8) & "|" & data(i, 9)
                If Not seen.Exists(address) Then
                    seen(address) = True
                    If result.Exists(key) Then a = result(key) Else a = Array(0&, 0&, 0&, 0&, 0&)
                    category = VD_AddressClass(CStr(data(i, 11)), CStr(data(i, 12)), volts, w, r)
                    n = 3
                    If category = "Write worse" Then n = 0
                    If category = "Read worse" Then n = 1
                    If category = "Equal" Then n = 2
                    a(n) = CLng(a(n)) + 1: a(4) = CLng(a(4)) + 1: result(key) = a
                End If
            End If
        Next i
        Next chunkStart
    End If
    Set VD_Stats = result
End Function

Private Function VD_Tendency(ByVal a As Variant) As String
    Dim kinds As Long
    If CLng(a(0)) > 0 Then kinds = kinds + 1
    If CLng(a(1)) > 0 Then kinds = kinds + 1
    If CLng(a(2)) > 0 Then kinds = kinds + 1
    If kinds = 0 Then VD_Tendency = "Insufficient": Exit Function
    If kinds > 1 Then VD_Tendency = "Mixed": Exit Function
    If CLng(a(0)) > 0 Then VD_Tendency = "Write worse"
    If CLng(a(1)) > 0 Then VD_Tendency = "Read worse"
    If CLng(a(2)) > 0 Then VD_Tendency = "Equal"
End Function

Private Function VD_Color(ByVal tendency As String) As Long
    Select Case tendency
        Case "Write worse": VD_Color = RGB(255, 217, 202)
        Case "Read worse": VD_Color = RGB(189, 215, 238)
        Case "Mixed": VD_Color = RGB(217, 210, 233)
        Case "Equal": VD_Color = RGB(226, 239, 218)
        Case Else: VD_Color = RGB(235, 235, 235)
    End Select
End Function

Private Function VD_FirstFail(ByVal counts As Object, ByVal scope As Collection, ByVal wt As String, ByVal dut As String, ByVal blk As Long) As Variant
    Dim item As Variant, key As String, missing As Boolean
    For Each item In scope
        key = wt & "|" & CStr(item(0)) & "|" & dut
        If Not counts.Exists(key) Then
            missing = True
        ElseIf Mid$(CStr(counts(key)), blk + 1, 1) = "1" Then
            If missing Then VD_FirstFail = Array(CStr(item(0)), Empty, "Earlier TNO missing; first failure uncertain") Else VD_FirstFail = Array(CStr(item(0)), item(1), IIf(IsEmpty(item(1)), "Function voltage missing", "First observed failure"))
            Exit Function
        End If
    Next item
    If missing Then VD_FirstFail = Array("", Empty, "Sweep incomplete") Else VD_FirstFail = Array("", Empty, "No recorded failure")
End Function

Private Function VD_BuildGrid(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal info As Variant) As Long
    Dim patterns As Object, wanted As Object, hard As Object, counts As Object, predecessors As Object, stats As Object
    Dim scope As Collection, idx As Worksheet, detail As Variant, first As Variant, a As Variant, packet(1 To 256, 1 To 16) As Variant
    Dim b As Long, r As Long, c As Long, j As Long, state As String, tendency As String, label As String, key As String
    VD_BuildGrid = 5
    If CStr(info(7)) <> "Fail BLK Count" Or Len(CStr(info(4))) = 0 Then Exit Function
    Set patterns = ReadStopPatterns(wb.Worksheets("Config"), wanted, hard)
    If Not patterns.Exists(CStr(info(4))) Then Exit Function
    Set counts = LoadMapCounts(wb.Worksheets("FBC by BLK"), wanted): Set predecessors = PatternPredecessors(patterns)
    Set scope = VD_Scope(wb.Worksheets("Config"), CStr(info(4)))
    Set stats = VD_Stats(wb, CStr(info(3)), CLng(info(5)), CLng(info(6)), CStr(info(4)))
    Set idx = wb.Worksheets("Import Index"): detail = patterns(CStr(info(4)))
    ws.Range("A5:P5").Merge: ws.Range("A5").Value2 = "BLK Voltage Distribution - " & CStr(info(4)) & " (double-click BLK)"
    SetMapDrillProperty ws, "ASCVoltageGrid", "1": SetMapDrillProperty ws, "ASCVoltageDetailTop", "0"
    idx.Range("BC1:BR257").ClearContents
    idx.Range("BC1:BR1").Value2 = Array("WT", "X", "Y", "Pattern", "BLK", "State", "First TNO", "First Voltage", "Observed tendency", "Write worse", "Read worse", "Equal", "Insufficient", "Captured", "Stop TNO", "Note")
    For b = 0 To 255
        r = 6 + b Mod 16: c = 1 + b \ 16: a = Array(0&, 0&, 0&, 0&, 0&)
        If stats.Exists(CStr(b)) Then a = stats(CStr(b))
        state = VD_State(counts, CStr(info(3)), info(5) & "|" & info(6), CStr(detail(0)), b, hard, wanted, predecessors)
        tendency = VD_Tendency(a): first = Array("", Empty, "")
        If state = "Fail" Then first = VD_FirstFail(counts, scope, CStr(info(3)), info(5) & "|" & info(6), b)
        label = "BLK" & b
        With ws.Cells(r, c)
            .Interior.Pattern = xlSolid: .Font.Color = vbBlack: .WrapText = True
            Select Case state
                Case "Fail"
                    label = label & vbLf & VD_VoltageText(first(1)) & vbLf
                    Select Case tendency
                        Case "Write worse": label = label & "W>R"
                        Case "Read worse": label = label & "R>W"
                        Case Else: label = label & tendency
                    End Select
                    If CLng(a(3)) > 0 And CLng(a(4)) > CLng(a(3)) Then label = label & "*"
                    .Interior.Color = VD_Color(tendency)
                Case "Pass": label = label & vbLf & "Pass": .Interior.Color = vbWhite
                Case "No Data": label = label & vbLf & "No Data": .Interior.Color = RGB(235, 235, 235)
                Case Else: label = label & vbLf & "Excluded": .Interior.Color = RGB(170, 170, 170)
            End Select
            .Value2 = label: .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(150, 150, 150)
        End With
        packet(b + 1, 1) = info(3): packet(b + 1, 2) = info(5): packet(b + 1, 3) = info(6): packet(b + 1, 4) = info(4)
        packet(b + 1, 5) = b: packet(b + 1, 6) = state: packet(b + 1, 7) = first(0): packet(b + 1, 8) = first(1): packet(b + 1, 9) = tendency
        For j = 0 To 4: packet(b + 1, j + 10) = a(j): Next j
        packet(b + 1, 15) = detail(0): packet(b + 1, 16) = first(2)
    Next b
    idx.Range("BC2:BR257").Value2 = packet
    ws.Range("A6:P21").RowHeight = 52
    ws.Range("A23:D23").Merge: ws.Range("A23").Value2 = "Write worse (observed)": ws.Range("A23:D23").Interior.Color = VD_Color("Write worse")
    ws.Range("E23:H23").Merge: ws.Range("E23").Value2 = "Read worse (observed)": ws.Range("E23:H23").Interior.Color = VD_Color("Read worse")
    ws.Range("I23:L23").Merge: ws.Range("I23").Value2 = "Mixed": ws.Range("I23:L23").Interior.Color = VD_Color("Mixed")
    ws.Range("M23:P23").Merge: ws.Range("M23").Value2 = "Equal": ws.Range("M23:P23").Interior.Color = VD_Color("Equal")
    ws.Range("A24:P24").Merge: ws.Range("A24").Value2 = "Captured Addresses only. * = some uncomparable captured bits. One-sided / missing Monitor evidence is Insufficient, not Pass."
    ws.Range("A25:P25").Merge: ws.Range("A25").Value2 = "Voltage = first Function FBC>0 in Config order, not full BLK Vmin. Color summarizes comparable observed Monitor voltages."
    ws.Rows("24:25").RowHeight = 32: ws.Range("A24:P25").WrapText = True
    VD_BuildGrid = 28
End Function

Public Function BLK_IsVoltageGridCell(ByVal ws As Object, ByVal target As Range) As Boolean
    If TypeName(ws) <> "Worksheet" Then Exit Function
    If ws.Name <> "Chip Fail Details" Or target.CountLarge <> 1 Then Exit Function
    If target.Row < 6 Or target.Row > 21 Or target.Column > 16 Or target.Column < 1 Then Exit Function
    BLK_IsVoltageGridCell = (VD_Property(ws, "ASCVoltageGrid") = "1" And Left$(CStr(ws.Range("A5").Value2), 24) = "BLK Voltage Distribution")
End Function

Private Function VD_ChipFBC(ByVal wb As Workbook, ByVal wt As String, ByVal x As Long, ByVal y As Long) As Object
    Dim result As Object, ws As Worksheet, data As Variant, values(0 To 255) As Variant
    Dim last As Long, first As Long, take As Long, i As Long, b As Long
    Set result = CreateObject("Scripting.Dictionary"): Set ws = wb.Worksheets("FBC by BLK"): last = PresentationLastRow(ws, 2)
    For first = 2 To last Step 512
        take = last - first + 1: If take > 512 Then take = 512
        data = ws.Cells(first, 1).Resize(take, 260).Value2
        For i = 1 To take
            If CStr(data(i, 1)) = wt And CLng(data(i, 3)) = x And CLng(data(i, 4)) = y Then
                For b = 0 To 255: values(b) = data(i, b + 5): Next b
                result(CStr(data(i, 2))) = values
            End If
        Next i
    Next first
    Set VD_ChipFBC = result
End Function
Public Function BLK_ShowBLKVoltage(ByVal wb As Workbook, ByVal ws As Object, ByVal target As Range) As String
    Dim idx As Worksheet, source As Worksheet, meta As Variant, data As Variant, values As Variant, header As Variant
    Dim patterns As Object, wanted As Object, hard As Object, counts As Object, predecessors As Object, chipFBC As Object, volts As Object
    Dim scope As Collection, voltageRows As New Collection, addressRows As New Collection, summary As New Collection
    Dim name As Variant, detail As Variant, item As Variant, xs As Collection, ys As Collection, series As Series, chart As ChartObject
    Dim b As Long, top As Long, previousTop As Long, last As Long, i As Long, j As Long, rowIndex As Long, chartTop As Long, addressesTop As Long
    Dim state As String, category As String, note As String, w As Variant, r As Variant, value As Variant, aX() As Variant, aY() As Variant
    Dim oldEvents As Boolean, oldScreen As Boolean, message As String
    Dim chunkStart As Long, chunkSize As Long
    Dim haveX As Boolean, xMin As Double, xMax As Double, xPad As Double
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    On Error GoTo Failed
    If Not BLK_IsVoltageGridCell(ws, target) Then Exit Function
    If InStr(1, CStr(ws.Range("A3").Value2), "updated", vbTextCompare) > 0 Then Err.Raise vbObjectError + 531, , "Refresh Map and Chip details after source changes."
    Application.EnableEvents = False: Application.ScreenUpdating = False
    b = (target.Row - 6) + 16 * (target.Column - 1)
    Set idx = wb.Worksheets("Import Index"): meta = idx.Cells(b + 2, 55).Resize(1, 16).Value2
    If CLng(meta(1, 5)) <> b Then Err.Raise vbObjectError + 531, , "BLK detail index changed; reopen Chip from Map."
    previousTop = CLng(Val(VD_Property(ws, "ASCVoltageDetailTop")))
    If previousTop > 0 Then
        For i = ws.ListObjects.Count To 1 Step -1
            If ws.ListObjects(i).Name = "tblBLKVoltage" Or ws.ListObjects(i).Name = "tblBLKObservedWR" Or ws.ListObjects(i).Name = "tblSelectedBLK" Then ws.ListObjects(i).Delete
        Next i
        last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        If last >= previousTop Then ws.Cells(previousTop, 1).Resize(last - previousTop + 1, 16).UnMerge: ws.Cells(previousTop, 1).Resize(last - previousTop + 1, 16).Clear
        top = previousTop
    Else
        top = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 4
    End If
    For i = ws.ChartObjects.Count To 1 Step -1
        If ws.ChartObjects(i).Name = "blkVoltageChart" Then ws.ChartObjects(i).Delete
    Next i
    SetMapDrillProperty ws, "ASCVoltageDetailTop", CStr(top)
    ws.Cells(top, 1).Resize(1, 16).Merge: ws.Cells(top, 1).Value2 = meta(1, 1) & " / X" & meta(1, 2) & " Y" & meta(1, 3) & " / BLK" & b & " / " & meta(1, 4)
    ws.Hyperlinks.Add ws.Cells(top + 1, 1), "", "'Chip Fail Details'!A5", , "Back to BLK grid"
    ws.Cells(top + 1, 4).Resize(1, 13).Merge: ws.Cells(top + 1, 4).Value2 = "Observed Monitor comparison only; not proof of full Address coverage."
    header = Array("WT", "X", "Y", "Pattern", "BLK", "State", "First TNO", "First Voltage", "Observed tendency", "Write worse", "Read worse", "Equal", "Insufficient", "Captured", "Stop TNO", "Note")
    Dim selected(0 To 15) As Variant
    For j = 1 To 16: selected(j - 1) = meta(1, j): Next j
    summary.Add selected
    rowIndex = PutCollection(ws, top + 3, "Selected BLK summary", header, summary, "tblSelectedBLK")
    ws.ListObjects("tblSelectedBLK").ListColumns(8).DataBodyRange.NumberFormat = "0.00#"
    Set patterns = ReadStopPatterns(wb.Worksheets("Config"), wanted, hard): Set counts = LoadMapCounts(wb.Worksheets("FBC by BLK"), wanted)
    Set predecessors = PatternPredecessors(patterns): Set chipFBC = VD_ChipFBC(wb, CStr(meta(1, 1)), CLng(meta(1, 2)), CLng(meta(1, 3)))
    For Each name In patterns
        detail = patterns(name): state = VD_State(counts, CStr(meta(1, 1)), meta(1, 2) & "|" & meta(1, 3), CStr(detail(0)), b, hard, wanted, predecessors)
        Set scope = VD_Scope(wb.Worksheets("Config"), CStr(name))
        For Each item In scope
            value = Empty: If chipFBC.Exists(CStr(item(0))) Then values = chipFBC(CStr(item(0))): value = values(b)
            voltageRows.Add Array(CStr(name), CStr(item(0)), CStr(item(2)), item(1), value, state)
        Next item
    Next name
    chartTop = rowIndex + 1
    rowIndex = PutCollection(ws, rowIndex, "Function FBC / voltage (Config sequence)", Array("Pattern", "TNO", "Func.", "Voltage (V)", "Raw FBC", "BLK eligibility"), voltageRows, "tblBLKVoltage")
    ws.ListObjects("tblBLKVoltage").ListColumns(4).DataBodyRange.NumberFormat = "0.00#"
    ws.ListObjects("tblBLKVoltage").ListColumns(5).DataBodyRange.NumberFormat = "0"
    Set chart = ws.ChartObjects.Add(ws.Cells(chartTop, 8).Left, ws.Cells(chartTop, 8).Top, ws.Range("H:P").Width, 280)
    chart.Name = "blkVoltageChart": chart.Chart.ChartType = xlXYScatter
    chart.Chart.ChartArea.Font.Name = "Calibri": chart.Chart.ChartArea.Font.Size = 11
    chart.Chart.HasTitle = True: chart.Chart.ChartTitle.Text = "BLK" & b & " Function FBC vs Voltage"
    For Each name In patterns
        Set xs = New Collection: Set ys = New Collection
        For Each item In voltageRows
            If CStr(item(0)) = CStr(name) And (CStr(item(5)) = "Pass" Or CStr(item(5)) = "Fail") Then
                If Not IsEmpty(item(3)) And Not IsEmpty(item(4)) Then
                    xs.Add CDbl(item(3)): ys.Add CDbl(item(4))
                    If Not haveX Then xMin = CDbl(item(3)): xMax = xMin: haveX = True
                    If CDbl(item(3)) < xMin Then xMin = CDbl(item(3))
                    If CDbl(item(3)) > xMax Then xMax = CDbl(item(3))
                End If
            End If
        Next item
        If xs.Count > 0 Then
            ReDim aX(1 To xs.Count): ReDim aY(1 To ys.Count)
            For j = 1 To xs.Count: aX(j) = xs(j): aY(j) = ys(j): Next j
            Set series = chart.Chart.SeriesCollection.NewSeries: series.Name = CStr(name): series.XValues = aX: series.Values = aY: series.MarkerSize = 6
        End If
    Next name
    If chart.Chart.SeriesCollection.Count > 0 Then
        chart.Chart.Axes(xlCategory).HasTitle = True: chart.Chart.Axes(xlCategory).AxisTitle.Text = "Function voltage (V)"
        chart.Chart.Axes(xlCategory).ReversePlotOrder = True
        xPad = (xMax - xMin) * 0.05: If xPad < 0.01 Then xPad = 0.05
        If xMin - xPad < 0 Then xPad = xMin
        chart.Chart.Axes(xlCategory).MinimumScale = xMin - xPad
        chart.Chart.Axes(xlCategory).MaximumScale = xMax + IIf(xPad > 0, xPad, 0.05)
        chart.Chart.Axes(xlCategory).TickLabels.NumberFormat = "0.00#"
        If xMax - xMin > 0.4 Then chart.Chart.Axes(xlCategory).MajorUnit = 0.1
        chart.Chart.Axes(xlValue).HasTitle = True: chart.Chart.Axes(xlValue).AxisTitle.Text = "FBC"
        chart.Chart.Axes(xlValue).MinimumScale = 0
        chart.Chart.HasLegend = True: chart.Chart.Legend.Position = xlLegendPositionBottom
    Else
        chart.Chart.ChartTitle.Text = "No eligible voltage / FBC points"
    End If
    ws.Cells(chartTop + 14, 8).Resize(3, 9).Merge: ws.Cells(chartTop + 14, 8).Value2 = "Excluded / missing points are not plotted. Raw FBC remains in the table. Repeated voltages may represent different TNOs."
    ws.Cells(chartTop + 14, 8).Resize(3, 9).WrapText = True
    addressesTop = rowIndex: If addressesTop < chartTop + 18 Then addressesTop = chartTop + 18
    Set source = wb.Worksheets("Monitor Fail Analysis"): last = source.Cells(source.Rows.Count, 1).End(xlUp).Row: Set volts = VD_MonitorVoltages(wb)
    If last >= 5 Then
        For chunkStart = 5 To last Step 512
        chunkSize = last - chunkStart + 1: If chunkSize > 512 Then chunkSize = 512
        data = source.Cells(chunkStart, 1).Resize(chunkSize, 16).Value2
        For i = 1 To UBound(data, 1)
            If CStr(data(i, 1)) = CStr(meta(1, 1)) And CStr(data(i, 2)) = CStr(meta(1, 4)) And CLng(data(i, 4)) = CLng(meta(1, 2)) And CLng(data(i, 5)) = CLng(meta(1, 3)) And CLng(data(i, 6)) = b Then
                category = VD_AddressClass(CStr(data(i, 11)), CStr(data(i, 12)), volts, w, r)
                addressRows.Add Array(data(i, 1), data(i, 2), data(i, 4), data(i, 5), data(i, 6), data(i, 7), data(i, 8), data(i, 9), w, r, category, data(i, 16))
            End If
        Next i
        Next chunkStart
    End If
    rowIndex = PutCollection(ws, addressesTop, "Captured Address W/R observed-voltage comparison", Array("WT", "Pattern", "X", "Y", "BLK", "WL", "BL", "IO", "Write max observed V", "Read max observed V", "Observed tendency", "Source Row"), addressRows, "tblBLKObservedWR")
    For i = 1 To addressRows.Count
        item = addressRows(i): ws.Hyperlinks.Add ws.Cells(addressesTop + 1 + i, 12), "", "'Fail Address'!A" & CLng(item(11)), , CStr(item(11))
    Next i
    If addressRows.Count > 0 Then ws.ListObjects("tblBLKObservedWR").ListColumns(9).DataBodyRange.NumberFormat = "0.00#": ws.ListObjects("tblBLKObservedWR").ListColumns(10).DataBodyRange.NumberFormat = "0.00#"
    With ws.Cells(top, 1).Resize(rowIndex - top, 16): .Font.Name = "Calibri": .Font.Size = 11: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: End With
    ws.Cells(top, 1).Resize(1, 16).Interior.Color = RGB(23, 50, 77): ws.Cells(top, 1).Resize(1, 16).Font.Color = vbWhite
    ws.Columns("A:P").ColumnWidth = 12
    chart.Left = ws.Cells(chartTop, 8).Left: chart.Width = ws.Range("H:P").Width
    ws.Activate: Application.Goto ws.Cells(top, 1), True
    BLK_ShowBLKVoltage = "BLK" & b & " details ready"
Done:
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    Exit Function
Failed:
    BLK_ShowBLKVoltage = "Error: " & Err.Description
    Resume Done
End Function
