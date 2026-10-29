Attribute VB_Name = "SplitThreeASCResults"
Option Explicit

' Independent second stage. Import into the completed ThreeASC result workbook.
' The original importer and Fail Address worksheet are never edited here.
Public Sub SplitWTFailAddresses()
    Static busy As Boolean
    Dim source As Worksheet, report As Workbook, targets(1 To 4) As Worksheet, summary As Worksheet
    Dim header As Variant, packet As Variant, output() As Variant, colWT() As Byte, classes() As Byte
    Dim noFail() As Boolean, chunks As Long
    Dim names As Variant, seen As Object, key As String, path As String, suffix As String
    Dim lastCol As Long, lastRow As Long, batch As Long, start As Long, count As Long
    Dim r As Long, c As Long, g As Long, rr As Long, category As Long, blockCounts(1 To 4) As Long
    Dim totals(1 To 4) As Long, written As Long, sourceCount As Long, zeroCount As Long, sequence As Long
    Dim screen As Boolean, events As Boolean, calc As XlCalculation, status As Variant
    Dim started As Date, stage As String, message As String, success As Boolean, lastCell As Range
    If busy Then Exit Sub
    busy = True
    screen = Application.ScreenUpdating: events = Application.EnableEvents
    calc = Application.Calculation: status = Application.StatusBar
    On Error GoTo Failed
    stage = "Validate original ThreeASC result"
    Set source = ThisWorkbook.Worksheets("Fail Address")
    lastCol = source.Cells(1, source.Columns.Count).End(xlToLeft).Column
    If lastCol < 9 Then Err.Raise vbObjectError + 101, , "Expected six address columns and three WT/TNO groups."
    header = source.Cells(1, 1).Resize(1, lastCol).Value2
    ValidateWTHeaders header, colWT
    Set lastCell = source.Cells.Find(What:="*", After:=source.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    lastRow = lastCell.Row
    started = Now
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Set seen = CreateObject("Scripting.Dictionary")
    Set report = Workbooks.Add(xlWBATWorksheet)
    Set summary = report.Worksheets(1): summary.Name = "Summary"
    names = Array("WT1 Fail", "WT2 Fail", "WT3 Fail", "WT Conflict")
    For g = 1 To 4
        Set targets(g) = report.Worksheets.Add(After:=report.Worksheets(report.Worksheets.Count))
        targets(g).Name = names(g - 1)
        targets(g).Cells(1, 1).Resize(1, lastCol).Value2 = header
    Next g
    targets(4).Cells(1, lastCol + 1).Value2 = "Conflict reason"
    batch = 131072 \ lastCol: If batch > 4096 Then batch = 4096
    For start = 2 To lastRow Step batch
        stage = "Read and classify source rows " & start
        count = lastRow - start + 1: If count > batch Then count = batch
        packet = source.Cells(start, 1).Resize(count, lastCol).Value2
        ReDim classes(1 To count)
        ReDim noFail(1 To count)
        For g = 1 To 4: blockCounts(g) = 0: Next g
        For r = 1 To count
            key = AddressKey(packet, r)
            If seen.Exists(key) Then Err.Raise vbObjectError + 102, , "Duplicate source address at row " & start + r - 1 & ": " & key
            seen.Add key, True
            category = ClassifyWTRow(packet, r, colWT)
            If category = 0 Then zeroCount = zeroCount + 1: noFail(r) = True: category = 4
            classes(r) = CByte(category)
            blockCounts(category) = blockCounts(category) + 1
            sourceCount = sourceCount + 1
        Next r
        For g = 1 To 4
            If blockCounts(g) > 0 Then
                ReDim output(1 To blockCounts(g), 1 To lastCol + IIf(g = 4, 1, 0))
                rr = 0
                For r = 1 To count
                    If classes(r) = g Then
                        rr = rr + 1
                        ' Copy every original cell in the same column; no reindexing or OR merge.
                        For c = 1 To lastCol: output(rr, c) = packet(r, c): Next c
                        If g = 4 Then
                            If noFail(r) Then
                                output(rr, lastCol + 1) = "No selected TNO fail in any WT"
                            Else
                                output(rr, lastCol + 1) = "Selected TNO fail in multiple WTs"
                            End If
                        End If
                    End If
                Next r
                targets(g).Cells(totals(g) + 2, 1).Resize(blockCounts(g), UBound(output, 2)).Value2 = output
                totals(g) = totals(g) + blockCounts(g)
            End If
        Next g
        chunks = chunks + 1
        If chunks Mod 8 = 0 Then
            Application.StatusBar = "Split WT results: " & sourceCount & " / " & lastRow - 1 & " addresses"
            DoEvents
        End If
    Next start
    stage = "Validate address totals"
    For g = 1 To 4: written = written + totals(g): Next g
    If written <> sourceCount Or sourceCount <> lastRow - 1 Then Err.Raise vbObjectError + 103, , "Address count reconciliation failed."
    stage = "Format split report"
    For g = 1 To 4
        FormatWTOutput targets(g), totals(g), lastCol, colWT, g = 4
    Next g
    summary.Range("A1:B1").Value2 = Array("ThreeASC second-stage split", "Value")
    summary.Range("A2:B12").Value2 = ArrayRows(ThisWorkbook.FullName, sourceCount, totals, zeroCount, DateDiff("s", started, Now))
    With summary.Range("A1:B12")
        .Font.Name = "Calibri": .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 28
    End With
    summary.Columns("A").ColumnWidth = 36: summary.Columns("B").ColumnWidth = 100
    summary.Range("B2").WrapText = True: summary.Rows(2).RowHeight = 48
    summary.Range("A1:B1").Interior.Color = RGB(23, 50, 77)
    summary.Range("A1:B1").Font.Color = vbWhite
    summary.Activate
    stage = "Save separate split report"
    If Len(ThisWorkbook.Path) = 0 Or InStr(1, ThisWorkbook.Path, "://", vbTextCompare) > 0 Then
        message = "Split complete. Save the new report manually; source workbook was not changed."
    Else
        path = ThisWorkbook.Path & Application.PathSeparator & Left$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & "_Split"
        Do
            suffix = "": If sequence > 0 Then suffix = "_" & Format$(sequence, "000")
            If Len(Dir$(path & suffix & ".xlsx")) = 0 Then Exit Do
            sequence = sequence + 1
        Loop
        If Len(path & suffix & ".xlsx") > 218 Then Err.Raise vbObjectError + 104, , "Report path exceeds Excel 2016 limit; shorten the source folder or filename."
        report.SaveAs Filename:=path & suffix & ".xlsx", FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        message = "Split complete. Saved as: " & report.FullName
    End If
    success = True
Cleanup:
    On Error Resume Next
    If Not success Then
        If Not report Is Nothing Then report.Close SaveChanges:=False
    End If
    Application.Calculation = calc: Application.EnableEvents = events
    Application.ScreenUpdating = screen: Application.StatusBar = status
    busy = False
    If Len(message) > 0 Then MsgBox message, IIf(success, vbInformation, vbExclamation)
    Exit Sub
Failed:
    message = "Split failed (" & Err.Number & "): " & Err.Description & vbCrLf & "Stage: " & stage & vbCrLf & "Original Fail Address worksheet was not changed."
    Resume Cleanup
End Sub

Private Sub ValidateWTHeaders(ByRef header As Variant, ByRef colWT() As Byte)
    Dim expected As Variant, seen(1 To 3) As Object, re As Object, matches As Object
    Dim c As Long, p As Long, tno As String, key As Variant
    expected = Array("X", "Y", "BLK", "WL", "BL", "IO")
    For c = 1 To 6
        If IsError(header(1, c)) Then Err.Raise vbObjectError + 105, , "Error in address header."
        If UCase$(Trim$(CStr(header(1, c)))) <> expected(c - 1) Then Err.Raise vbObjectError + 105, , "Address headers must be X/Y/BLK/WL/BL/IO."
    Next c
    For p = 1 To 3: Set seen(p) = CreateObject("Scripting.Dictionary"): Next p
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True: re.Pattern = "^WT([123])\s+([0-9]{3})$"
    ReDim colWT(1 To UBound(header, 2))
    For c = 7 To UBound(header, 2)
        If IsError(header(1, c)) Then Err.Raise vbObjectError + 106, , "Error in WT/TNO header."
        Set matches = re.Execute(Trim$(CStr(header(1, c))))
        If matches.Count <> 1 Then Err.Raise vbObjectError + 106, , "Expected ThreeASC WT + newline + TNO at column " & c
        p = CLng(matches(0).SubMatches(0)): tno = CStr(matches(0).SubMatches(1))
        If CLng(tno) < 100 Then Err.Raise vbObjectError + 106, , "TNO must be 100-999."
        If seen(p).Exists(tno) Then Err.Raise vbObjectError + 107, , "Duplicate WT/TNO header: WT" & p & " / " & tno
        seen(p).Add tno, True: colWT(c) = CByte(p)
    Next c
    If seen(1).Count = 0 Or seen(1).Count <> seen(2).Count Or seen(1).Count <> seen(3).Count Then Err.Raise vbObjectError + 108, , "WT1/2/3 must contain the same nonempty TNO set."
    For Each key In seen(1)
        If Not seen(2).Exists(key) Or Not seen(3).Exists(key) Then Err.Raise vbObjectError + 108, , "WT1/2/3 TNO sets differ."
    Next key
End Sub

Private Function ClassifyWTRow(ByRef packet As Variant, ByVal r As Long, ByRef colWT() As Byte) As Long
    Dim found(1 To 3) As Boolean, c As Long, p As Long, count As Long, result As Long, text As String
    For c = 7 To UBound(packet, 2)
        If Not IsEmpty(packet(r, c)) Then
            If VarType(packet(r, c)) = vbDouble Then
                If packet(r, c) <> 1 Then Err.Raise vbObjectError + 109, , "Expected blank or 1 in source TNO column " & c
                found(colWT(c)) = True
            Else
                If IsError(packet(r, c)) Then Err.Raise vbObjectError + 109, , "Error value in source TNO column " & c
                text = Trim$(CStr(packet(r, c)))
                If Len(text) > 0 Then
                    If text <> "1" Or VarType(packet(r, c)) = vbBoolean Then Err.Raise vbObjectError + 109, , "Expected blank or 1 in source TNO column " & c
                    found(colWT(c)) = True
                End If
            End If
        End If
    Next c
    For p = 1 To 3
        If found(p) Then count = count + 1: result = p
    Next p
    If count > 1 Then result = 4
    ClassifyWTRow = result
End Function

Private Function AddressKey(ByRef packet As Variant, ByVal r As Long) As String
    Dim c As Long, text As String, n As Double, key As String
    For c = 1 To 6
        If IsError(packet(r, c)) Then Err.Raise vbObjectError + 110, , "Error in source address."
        text = Trim$(CStr(packet(r, c)))
        If Len(text) = 0 Or Not IsNumeric(text) Then Err.Raise vbObjectError + 110, , "Address must have six integers."
        n = CDbl(text)
        If n <> Fix(n) Or n < -2147483648# Or n > 2147483647# Then Err.Raise vbObjectError + 110, , "Invalid source address integer."
        If c > 1 Then key = key & ","
        key = key & CStr(CLng(n))
    Next c
    AddressKey = key
End Function

Private Sub FormatWTOutput(ByVal ws As Worksheet, ByVal count As Long, ByVal lastCol As Long, ByRef colWT() As Byte, ByVal conflict As Boolean)
    Dim c As Long, color As Long, endCol As Long, first As Long, phase As Long
    endCol = lastCol: If conflict Then endCol = lastCol + 1
    With ws.Cells(1, 1).Resize(count + 1, endCol)
        .Font.Name = "Calibri": .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 26
        .AutoFilter
    End With
    ws.Columns("A:F").ColumnWidth = 10
    With ws.Cells(1, 1).Resize(1, endCol)
        .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: .RowHeight = 60
    End With
    c = 7
    Do While c <= lastCol
        first = c: phase = colWT(c)
        Do While c <= lastCol
            If colWT(c) <> phase Then Exit Do
            c = c + 1
        Loop
        Select Case phase
            Case 1: color = RGB(219, 234, 254)
            Case 2: color = RGB(255, 237, 213)
            Case 3: color = RGB(220, 252, 231)
        End Select
        With ws.Cells(1, first).Resize(1, c - first)
            .Interior.Color = color: .Font.Color = RGB(23, 50, 77)
            .Orientation = 90: .WrapText = True: .ColumnWidth = 6
        End With
        If count > 0 Then
            With ws.Cells(2, first).Resize(count, c - first).FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=1")
                .Interior.Color = color
            End With
        End If
    Loop
    If conflict Then ws.Columns(endCol).ColumnWidth = 44
End Sub

Private Function ArrayRows(ByVal sourcePath As String, ByVal count As Long, ByRef totals() As Long, ByVal zeroCount As Long, ByVal seconds As Long) As Variant
    Dim data(1 To 11, 1 To 2) As Variant
    data(1, 1) = "Source workbook": data(1, 2) = sourcePath
    data(2, 1) = "Source address count": data(2, 2) = count
    data(3, 1) = "WT1 Fail": data(3, 2) = totals(1)
    data(4, 1) = "WT2 Fail": data(4, 2) = totals(2)
    data(5, 1) = "WT3 Fail": data(5, 2) = totals(3)
    data(6, 1) = "WT Conflict": data(6, 2) = totals(4)
    data(7, 1) = "Conflict: no selected fail": data(7, 2) = zeroCount
    data(8, 1) = "Conflict: multiple WTs": data(8, 2) = totals(4) - zeroCount
    data(9, 1) = "Count reconciliation": data(9, 2) = "PASS"
    data(10, 1) = "Run time seconds": data(10, 2) = seconds
    data(11, 1) = "Classification unit": data(11, 2) = "X/Y/BLK/WL/BL/IO; source WT/TNO headers"
    ArrayRows = data
End Function
