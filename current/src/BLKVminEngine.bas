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


Public Function BLK_BuildFailMaps(ByVal wb As Workbook) As String
    Dim patterns As Object, wanted As Object, counts As Object, coords As Object, errors As Object, used As Object
    Dim phaseCount As Long, includedMask As Variant, maximumCount As Long
    Dim priorPatterns As Object, categoryByWT(1 To 3) As Object, chipCat As Long
    Dim hard As Object
    Dim locations As Worksheet, summary As Worksheet, map As Worksheet, ws As Worksheet
    Dim data As Variant, item As Variant, name As Variant, detail As Variant, xy As Variant, ordered As Collection
    Dim i As Long, phase As Long, last As Long, minX As Long, maxX As Long, minY As Long, maxY As Long
    Dim width As Long, col As Long, row As Long, origin As Long, sheetName As String, result As Variant
    Dim output() As Variant, n As Long, total As Double, text As String, existing As Boolean
    On Error GoTo Failed
    BLK_InvalidateCoverageMap wb
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
                If chipCat = 1 Or chipCat = 2 Then result = 0
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
    Set ws = wb.Worksheets("Fail BLK Address Map")
    On Error GoTo 0
    If Not ws Is Nothing Then ws.UsedRange.UnMerge: ws.UsedRange.Clear
End Sub

Public Sub BLK_BuildCoverageMap(ByVal wb As Workbook)
    'Right cell counts distinct BLKs with Missing Count > 0 (not the missing-bit sum).
    Dim ws As Worksheet, source As Worksheet, report As Worksheet
    Dim summaries As Variant, data As Variant, names As Object, coords As Object, losses As Object, seen As Object
    Dim cats(1 To 3) As Object, ordered As Collection, item As Variant, name As Variant, xy As Variant
    Dim i As Long, j As Long, phase As Long, phaseCount As Long, last As Long, first As Long, take As Long
    Dim minX As Long, maxX As Long, minY As Long, maxY As Long, x As Long, y As Long
    Dim top As Long, height As Long, panelWidth As Long, origin As Long, row As Long, col As Long
    Dim key As String, blockKey As String, chip As String, status As String, cat As Long, loss As Long
    Dim fail As Long, maximum As Long, cells As Range, colorRule As Object
    Set source = wb.Worksheets("BLK Fail Count"): last = source.Cells(source.Rows.Count, 1).End(xlUp).Row
    If last < 2 Then Err.Raise vbObjectError + 492, , "Build Fail Maps before the Address Map."
    summaries = source.Range("A2:I" & last).Value2
    Set names = CreateObject("Scripting.Dictionary"): Set coords = CreateObject("Scripting.Dictionary")
    Set losses = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary")
    phaseCount = BLK_PhaseCount(wb)
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
    For first = 5 To last Step 2048
        take = last - first + 1: If take > 2048 Then take = 2048
        data = report.Cells(first, 1).Resize(take, 12).Value2
        For i = 1 To UBound(data, 1)
            If CDbl(data(i, 11)) > 0 Then
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
    Set ws = wb.Worksheets("Fail BLK Address Map")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=source): ws.Name = "Fail BLK Address Map"
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
                If cat <> 1 And cat <> 2 Then
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
    If wb.Worksheets("Log").Index <> wb.Worksheets.Count Then wb.Worksheets("Log").Move After:=wb.Worksheets(wb.Worksheets.Count)
    If wb.Worksheets("Chip Locations").Index <> wb.Worksheets.Count - 1 Then wb.Worksheets("Chip Locations").Move Before:=wb.Worksheets("Log")
End Sub

Public Sub BLK_ApplyPresentation(ByVal wb As Workbook)
    Dim previousWindow As Window, previousSheet As Object, window As Window, ws As Worksheet, lastRow As Long, lastCol As Long
    On Error GoTo Restore
    Set previousWindow = Application.ActiveWindow: Set previousSheet = Application.ActiveSheet
    wb.Worksheets("Config").Range("H1:N901").Borders.LineStyle = xlNone
    If wb.Worksheets("Log").Index <> wb.Worksheets.Count Then wb.Worksheets("Log").Move After:=wb.Worksheets(wb.Worksheets.Count)
    If wb.Worksheets("Chip Locations").Index <> wb.Worksheets.Count - 1 Then wb.Worksheets("Chip Locations").Move Before:=wb.Worksheets("Log")
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
                If ws.Name = "Fail Address" Then lastRow = PresentationLastRow(ws, 1)
                If lastRow < 1 Then lastRow = 1
                If ws.Name <> "Config" Then
                    ws.Cells(1, 1).Resize(1, lastCol).Borders.LineStyle = xlContinuous
                    ws.Cells(1, 1).Resize(lastRow, lastCol).BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(80, 80, 80)
                End If
            End If
        Next ws
    Next window
    wb.Worksheets("Config").Range("H1:N901").Borders.LineStyle = xlNone
    If wb.Worksheets("Log").Index <> wb.Worksheets.Count Then wb.Worksheets("Log").Move After:=wb.Worksheets(wb.Worksheets.Count)
    If wb.Worksheets("Chip Locations").Index <> wb.Worksheets.Count - 1 Then wb.Worksheets("Chip Locations").Move Before:=wb.Worksheets("Log")
Restore:
    On Error Resume Next
    If Not previousWindow Is Nothing Then previousWindow.Activate
    If Not previousSheet Is Nothing Then previousSheet.Activate
End Sub

Public Sub BLK_SizeOutput(ByVal ws As Worksheet)
    Dim lastRow As Long, lastCol As Long, c As Long, firstData As Long, height As Double, maximum As Double
    If ws.Name = "Config" Or ws.Name = "Import Index" Or ws.Name = "FBC by BLK" Then Exit Sub
    If Right$(ws.Name, 8) = "Fail Map" Or ws.Name = "Fail BLK Address Map" Then Exit Sub
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
