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
    groupSize = UBound(tnos) + 1
    If CDbl(dataRow) + CDbl(duts.Count) * groupSize - 1 > ws.Rows.Count Then Err.Raise vbObjectError + 404, , "Test Name sheet exceeds row limit."
    endRow = dataRow
    For Each dut In duts
        group = BuildDUTGroup(mData, tnos, volts, CStr(dut), mThreshold)
        WriteGroup ws, endRow, group: endRow = endRow + groupSize
    Next dut
    With ws.Cells(dataRow, 1).Resize(endRow - dataRow, 260)
        .Font.Name = "Calibri": .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 26
    End With
    For i = 1 To duts.Count
        block = dataRow + i * groupSize - 1
        ws.Cells(block, 1).Resize(1, 4).Merge
        With ws.Cells(block, 1).Resize(1, 260)
            .Interior.Color = RGB(37, 99, 168): .Font.Color = vbWhite: .Font.Bold = True
        End With
    Next i
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
