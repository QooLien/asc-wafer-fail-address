Attribute VB_Name = "SplitWTBySum"
Option Explicit

' Stage two only: independent WT sums; an address may appear on multiple Fail sheets.
Public Sub SplitWTByTNOSum()
    Static busy As Boolean
    Dim src As Worksheet, report As Workbook, statusSheet As Worksheet, dst(1 To 3) As Worksheet
    Dim header As Variant, packet As Variant, statusData() As Variant, output() As Variant
    Dim colWT() As Byte, colMap() As Long, counts(1 To 3) As Long, sums() As Long
    Dim totals(1 To 3) As Long, marks(1 To 3) As Double, blockCount(1 To 3) As Long
    Dim seen As Object, lastCell As Range, key As String, path As String, candidate As String, fso As Object
    Dim start As Long, count As Long, lastRow As Long, lastCol As Long, n As Long, batch As Long
    Dim r As Long, c As Long, p As Long, rr As Long, processed As Long, sequence As Long, chunks As Long
    Dim screen As Boolean, events As Boolean, calc As XlCalculation, oldStatus As Variant
    Dim started As Date, stage As String, message As String, success As Boolean, color As Long
    Dim outHeader() As Variant, stats(1 To 3, 1 To 4) As Variant
    If busy Then Exit Sub
    busy = True
    screen = Application.ScreenUpdating: events = Application.EnableEvents
    calc = Application.Calculation: oldStatus = Application.StatusBar
    On Error GoTo Failed
    stage = "Validate original ThreeASC headers"
    Set src = ThisWorkbook.Worksheets("Fail Address")
    lastCol = src.Cells(1, src.Columns.Count).End(xlToLeft).Column
    If lastCol < 9 Then Err.Raise vbObjectError + 201, , "Expected ThreeASC WT1/2/3 headers."
    header = src.Cells(1, 1).Resize(1, lastCol).Value2
    ValidateWTHeaders header, colWT
    n = (lastCol - 6) \ 3: ReDim colMap(1 To 3, 1 To n)
    For c = 7 To lastCol
        p = colWT(c): counts(p) = counts(p) + 1: colMap(p, counts(p)) = c
    Next c
    Set lastCell = src.Cells.Find(What:="*", After:=src.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    lastRow = lastCell.Row: started = Now
    Application.ScreenUpdating = False: Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Set seen = CreateObject("Scripting.Dictionary")
    Set report = Workbooks.Add(xlWBATWorksheet)
    Set statusSheet = report.Worksheets(1): statusSheet.Name = "WT Status"
    statusSheet.Range("A1:L1").Value2 = Array("X", "Y", "BLK", "WL", "BL", "IO", "WT1 Sum", "WT1 Status", "WT2 Sum", "WT2 Status", "WT3 Sum", "WT3 Status")
    For p = 1 To 3
        Set dst(p) = report.Worksheets.Add(After:=report.Worksheets(report.Worksheets.Count))
        dst(p).Name = "WT" & p & " Fail"
        ReDim outHeader(1 To 1, 1 To n + 8)
        For c = 1 To 6: outHeader(1, c) = header(1, c): Next c
        For c = 1 To n: outHeader(1, 6 + c) = header(1, colMap(p, c)): Next c
        outHeader(1, n + 7) = "WT" & p & " Sum": outHeader(1, n + 8) = "Status"
        dst(p).Cells(1, 1).Resize(1, n + 8).Value2 = outHeader
    Next p
    batch = 131072 \ lastCol: If batch > 4096 Then batch = 4096
    For start = 2 To lastRow Step batch
        stage = "Sum WT TNO values at source row " & start
        count = lastRow - start + 1: If count > batch Then count = batch
        packet = src.Cells(start, 1).Resize(count, lastCol).Value2
        ReDim sums(1 To count, 1 To 3): ReDim statusData(1 To count, 1 To 12)
        For p = 1 To 3: blockCount(p) = 0: Next p
        For r = 1 To count
            key = AddressKey(packet, r)
            If seen.Exists(key) Then Err.Raise vbObjectError + 202, , "Duplicate source address at row " & start + r - 1 & ": " & key
            seen.Add key, True
            SumWTRow packet, r, colWT, sums
            For c = 1 To 6: statusData(r, c) = packet(r, c): Next c
            For p = 1 To 3
                statusData(r, 5 + 2 * p) = sums(r, p)
                statusData(r, 6 + 2 * p) = IIf(sums(r, p) = 0, "Pass", "Fail")
                If sums(r, p) > 0 Then blockCount(p) = blockCount(p) + 1
                marks(p) = marks(p) + sums(r, p)
            Next p
        Next r
        statusSheet.Cells(processed + 2, 1).Resize(count, 12).Value2 = statusData
        For p = 1 To 3
            If blockCount(p) > 0 Then
                ReDim output(1 To blockCount(p), 1 To n + 8)
                BuildWTFailBatch packet, sums, colMap, p, n, output, rr
                If rr <> blockCount(p) Then Err.Raise vbObjectError + 203, , "WT Fail count reconciliation failed."
                dst(p).Cells(totals(p) + 2, 1).Resize(rr, n + 8).Value2 = output
                totals(p) = totals(p) + rr
            End If
        Next p
        processed = processed + count: chunks = chunks + 1
        If chunks Mod 8 = 0 Then
            Application.StatusBar = "Independent WT sums: " & processed & " / " & lastRow - 1
            DoEvents
        End If
    Next start
    If processed <> lastRow - 1 Then Err.Raise vbObjectError + 204, , "Source address count reconciliation failed."
    stage = "Format independent WT report"
    For p = 1 To 3
        Select Case p
            Case 1: color = RGB(219, 234, 254)
            Case 2: color = RGB(255, 237, 213)
            Case 3: color = RGB(220, 252, 231)
        End Select
        FormatSumSheet dst(p), totals(p), n + 8
        With dst(p).Cells(1, 7).Resize(1, n)
            .Interior.Color = color: .Font.Color = RGB(23, 50, 77)
            .Orientation = 90: .WrapText = True: .ColumnWidth = 6
        End With
        If totals(p) > 0 Then
            With dst(p).Cells(2, 7).Resize(totals(p), n).FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=1")
                .Interior.Color = color
            End With
        End If
        stats(p, 1) = "WT" & p: stats(p, 2) = processed - totals(p)
        stats(p, 3) = totals(p): stats(p, 4) = marks(p)
    Next p
    FormatSumSheet statusSheet, processed, 12
    statusSheet.Rows(1).RowHeight = 40
    statusSheet.Range("N1:Q1").Value2 = Array("WT", "Pass addresses", "Fail addresses", "Fail TNO marks")
    statusSheet.Range("N2:Q4").Value2 = stats
    statusSheet.Range("N6:O8").Value2 = ArrayInfo(ThisWorkbook.FullName, processed, DateDiff("s", started, Now))
    With statusSheet.Range("N1:Q8")
        .Font.Name = "Calibri": .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .ColumnWidth = 22
    End With
    statusSheet.Activate
    stage = "Save separate WT Sum report"
    If Len(ThisWorkbook.Path) = 0 Or InStr(1, ThisWorkbook.Path, "://", vbTextCompare) > 0 Then
        message = "WT Sum report complete. Save the new report manually."
    Else
        Set fso = CreateObject("Scripting.FileSystemObject")
        path = ThisWorkbook.Path & Application.PathSeparator & Left$(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & "_WT_Sum"
        candidate = path & ".xlsx"
        Do While fso.FileExists(candidate)
            sequence = sequence + 1: candidate = path & "_" & Format$(sequence, "000") & ".xlsx"
        Loop
        If Len(candidate) > 218 Then Err.Raise vbObjectError + 205, , "Report path is too long for Excel 2016."
        report.SaveAs Filename:=candidate, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
        message = "WT Sum report complete: " & report.FullName
    End If
    success = True
Cleanup:
    On Error Resume Next
    If Not success Then
        If Not report Is Nothing Then report.Close SaveChanges:=False
    End If
    Application.Calculation = calc: Application.EnableEvents = events
    Application.ScreenUpdating = screen: Application.StatusBar = oldStatus: busy = False
    If Len(message) > 0 Then MsgBox message, IIf(success, vbInformation, vbExclamation)
    Exit Sub
Failed:
    message = "WT Sum failed (" & Err.Number & "): " & Err.Description & vbCrLf & "Stage: " & stage & vbCrLf & "Original Fail Address was not changed."
    Resume Cleanup
End Sub

Private Sub SumWTRow(ByRef packet As Variant, ByVal r As Long, ByRef colWT() As Byte, ByRef sums() As Long)
    Dim c As Long, text As String
    For c = 7 To UBound(packet, 2)
        If Not IsEmpty(packet(r, c)) Then
            If VarType(packet(r, c)) = vbDouble Then
                If packet(r, c) <> 1 Then Err.Raise vbObjectError + 206, , "Expected blank or 1 in TNO column " & c
                sums(r, colWT(c)) = sums(r, colWT(c)) + 1
            Else
                If IsError(packet(r, c)) Then Err.Raise vbObjectError + 206, , "Error in TNO column " & c
                text = Trim$(CStr(packet(r, c)))
                If Len(text) > 0 Then
                    If text <> "1" Or VarType(packet(r, c)) = vbBoolean Then Err.Raise vbObjectError + 206, , "Expected blank or 1 in TNO column " & c
                    sums(r, colWT(c)) = sums(r, colWT(c)) + 1
                End If
            End If
        End If
    Next c
End Sub

Private Sub BuildWTFailBatch(ByRef packet As Variant, ByRef sums() As Long, ByRef colMap() As Long, ByVal phase As Long, ByVal n As Long, ByRef output() As Variant, ByRef rows As Long)
    Dim r As Long, c As Long
    rows = 0
    For r = 1 To UBound(packet, 1)
        If sums(r, phase) > 0 Then
            rows = rows + 1
            For c = 1 To 6: output(rows, c) = packet(r, c): Next c
            For c = 1 To n: output(rows, 6 + c) = packet(r, colMap(phase, c)): Next c
            output(rows, n + 7) = sums(r, phase): output(rows, n + 8) = "Fail"
        End If
    Next r
End Sub

Private Sub FormatSumSheet(ByVal ws As Worksheet, ByVal count As Long, ByVal columns As Long)
    With ws.Cells(1, 1).Resize(count + 1, columns)
        .Font.Name = "Calibri": .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 26
        .AutoFilter: .ColumnWidth = 12
    End With
    With ws.Cells(1, 1).Resize(1, columns)
        .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: .RowHeight = 60
    End With
End Sub

Private Function ArrayInfo(ByVal path As String, ByVal count As Long, ByVal seconds As Long) As Variant
    Dim data(1 To 3, 1 To 2) As Variant
    data(1, 1) = "Source workbook": data(1, 2) = path
    data(2, 1) = "Address count": data(2, 2) = count
    data(3, 1) = "Run time seconds": data(3, 2) = seconds
    ArrayInfo = data
End Function

' Shared source-header and address validators follow, unchanged from the stage-two splitter.
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

