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
    Dim color As Long, note As String
    GroupSequence mSequence, mVoltage, mTestNames, funcName, tnos, volts
    n = UBound(tnos)
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
    If CDbl(ordered.Count) * (n + 1) * 3 + 1 > ws.Rows.Count Then Err.Raise vbObjectError + 405, , "Comparison sheet exceeds row limit."
    header(1, 1) = "TNO / Vmin Label": header(1, 2) = "Func.": header(1, 3) = "Test Voltage"
    header(1, 4) = "X": header(1, 5) = "Y": header(1, 6) = "WT"
    For b = 0 To 255: header(1, b + 7) = "BLK " & b: Next b
    ws.UsedRange.UnMerge: ws.UsedRange.Clear
    ws.Cells(1, 1).Resize(1, 262).Value2 = header
    dest = 2
    For Each dut In ordered
        out = BuildComparisonGroup(source, rows, CStr(dut), funcName, tnos, volts)
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
        With ws.Cells(dest + rr - 3, 1).Resize(3, 262)
            .Interior.Color = RGB(37, 99, 168): .Font.Color = vbWhite: .Font.Bold = True
        End With
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

Private Function BuildComparisonGroup(ByRef source As Variant, ByVal rows As Object, ByVal dut As String, ByVal funcName As String, ByRef tnos() As String, ByRef volts() As Double) As Variant
    Dim out() As Variant, xy As Variant, n As Long, rr As Long, i As Long, phase As Long, key As String, src As Long, b As Long
    n = UBound(tnos)
        xy = Split(dut, "|")
        ReDim out(1 To (n + 1) * 3, 1 To 262)
        rr = 0
        For i = 1 To n + 1
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
