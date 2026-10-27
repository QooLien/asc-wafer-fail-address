Attribute VB_Name = "ImportFailAddresses"
Option Explicit

Private Const READ_BYTES As Long = 65536
Private Const MAX_PENDING_RECORD As Long = 16384
Private mIntegerRegex As Object

' Embedded macro: bounded file buffers, batched log/output, one row index per address.
Public Sub ImportASCFailAddresses()
    Static busy As Boolean
    Dim cfg As Worksheet, dst As Worksheet, work As Worksheet
    Dim tnos As Object, addresses As Object, logs As Object, records As Object
    Dim fd As Object, fields As Variant, record As Variant, v As Variant
    Dim raw As String, pendingCR As String, buffer As String, encoding As String
    Dim path As String, runID As String, stage As String, headerTNO As String, s As String
    Dim r As Long, j As Long, targetRow As Long, sourceLine As Long, selectedSections As Long
    Dim skipped As Long, accepted As Long, uniqueCount As Long, chunkCount As Long
    Dim fileNo As Integer, fileOpen As Boolean, fileSize As Long, batchRows As Long
    Dim n As Double, oldScreen As Boolean, oldAlerts As Boolean, oldStatus As Variant
    Dim errorNumber As Long, errorText As String, logError As String
    If busy Then MsgBox "An import is already running.", vbExclamation: Exit Sub
    busy = True
    On Error GoTo Failed
    oldScreen = Application.ScreenUpdating
    oldAlerts = Application.DisplayAlerts
    oldStatus = Application.StatusBar
    runID = Format$(Now, "yyyymmdd_hhnnss") & "_" & Format$(Timer * 100, "00000000")
    Set logs = CreateObject("Scripting.Dictionary")
    Set records = CreateObject("Scripting.Dictionary")
    Set addresses = CreateObject("Scripting.Dictionary")
    Set tnos = CreateObject("Scripting.Dictionary")
    stage = "Validate Config"
    AddImportLog logs, runID, "INFO", stage, "", "", 0, 0, "Streaming import started; chunk bytes=65536.", ""
    Set cfg = ThisWorkbook.Worksheets("Config")
    Set dst = ThisWorkbook.Worksheets("Fail Addresses")
    For r = 2 To 901
        v = cfg.Cells(r, 1).Value2
        If IsError(v) Then Err.Raise vbObjectError + 1, , "Invalid TNO at Config!A" & r
        s = CleanSpace(CStr(v))
        If Len(s) > 0 Then
            If Not IntegerInRange(s, 100, 999, n) Or Len(s) <> 3 Then Err.Raise vbObjectError + 2, , "TNO must be 100-999 at Config!A" & r
            If Not tnos.Exists(s) Then tnos.Add s, tnos.Count + 7
        End If
    Next r
    If tnos.Count = 0 Then Err.Raise vbObjectError + 3, , "Enter TNOs in Config!A2:A901."
    stage = "Select ASC"
    Set fd = Application.FileDialog(3)
    With fd
        .Title = "Select ASC raw data"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "ASC raw data", "*.asc"
        If .Show <> -1 Then
            AddImportLog logs, runID, "INFO", stage, "", "", 0, 0, "User cancelled; previous results retained.", ""
            FlushImportLog logs
            GoTo Done
        End If
        path = .SelectedItems(1)
    End With
    stage = "Open ASC"
    If CreateObject("Scripting.FileSystemObject").GetFile(path).Size > 2147483647# Then Err.Raise vbObjectError + 8, , "ASC exceeds the VBA binary-reader 2 GB limit. Split the input file."
    fileNo = FreeFile
    Open path For Binary Access Read As #fileNo
    fileOpen = True
    fileSize = LOF(fileNo)
    DetectEncoding fileNo, encoding
    AddImportLog logs, runID, "INFO", stage, path, "", 0, 0, "File bytes=" & fileSize & "; encoding=" & encoding & "; bounded streaming enabled.", ""
    FlushImportLog logs
    Application.ScreenUpdating = False
    Set work = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    work.Name = "ASC_Work_" & Format$(Timer * 100, "00000000")
    fields = Array("X", "Y", "BLK", "WL", "BL", "IO")
    work.Range("A1:F1").Value2 = fields
    For Each v In tnos.Keys: work.Cells(1, CLng(tnos(v))).Value2 = CLng(v): Next v
    sourceLine = 1
    Do While Seek(fileNo) <= fileSize
        stage = "Read ASC chunk"
        raw = pendingCR & ReadASCChunk(fileNo, encoding)
        pendingCR = ""
        If Right$(raw, 1) = vbCr Then pendingCR = vbCr: raw = Left$(raw, Len(raw) - 1)
        raw = Replace(Replace(raw, vbCrLf, vbLf), vbCr, vbLf)
        buffer = buffer & raw
        raw = ""
        stage = "Parse addresses"
        ParseASCBuffer buffer, False, headerTNO, sourceLine, selectedSections, skipped, tnos, logs, runID, path, records
        stage = "Stage address results"
        For Each v In records.Keys
            record = records(v)
            If Not addresses.Exists(record(0)) Then
                If uniqueCount >= dst.Rows.Count - 1 Then Err.Raise vbObjectError + 4, , "Too many unique addresses for one Excel sheet."
                uniqueCount = uniqueCount + 1
                targetRow = uniqueCount + 1
                addresses.Add record(0), targetRow
                work.Cells(targetRow, 1).Resize(1, 6).Value2 = Array(record(2), record(3), record(4), record(5), record(6), record(7))
            Else
                targetRow = CLng(addresses(record(0)))
            End If
            work.Cells(targetRow, CLng(tnos(record(1)))).Value2 = 1
        Next v
        accepted = accepted + records.Count
        records.RemoveAll
        If logs.Count >= 500 Then FlushImportLog logs
        chunkCount = chunkCount + 1
        If chunkCount Mod 16 = 0 Then
            Application.StatusBar = "ASC: " & Format$((Seek(fileNo) - 1) / 1024, "#,##0") & " / " & Format$(fileSize / 1024, "#,##0") & " KB; unique=" & uniqueCount & "; skipped=" & skipped
            DoEvents
        End If
    Loop
    Close #fileNo
    fileOpen = False
    If Len(pendingCR) > 0 Then buffer = buffer & vbLf
    stage = "Parse final buffer"
    ParseASCBuffer buffer, True, headerTNO, sourceLine, selectedSections, skipped, tnos, logs, runID, path, records
    ' At EOF only a final TNO or unfinished bracket can remain; complete records
    ' are already handled in the chunk loop, but stage any valid final records too.
    For Each v In records.Keys
        record = records(v)
        If Not addresses.Exists(record(0)) Then
            If uniqueCount >= dst.Rows.Count - 1 Then Err.Raise vbObjectError + 4, , "Too many unique addresses for one Excel sheet."
            uniqueCount = uniqueCount + 1: targetRow = uniqueCount + 1
            addresses.Add record(0), targetRow
            work.Cells(targetRow, 1).Resize(1, 6).Value2 = Array(record(2), record(3), record(4), record(5), record(6), record(7))
        Else
            targetRow = CLng(addresses(record(0)))
        End If
        work.Cells(targetRow, CLng(tnos(record(1)))).Value2 = 1
    Next v
    accepted = accepted + records.Count
    If selectedSections = 0 Then Err.Raise vbObjectError + 5, , "No selected tno= section found. Check TNO and file encoding."
    If accepted = 0 And skipped > 0 Then Err.Raise vbObjectError + 6, , "All selected address records were rejected. See Import Log; previous results retained."
    stage = "Write output batches"
    dst.UsedRange.ClearContents
    For r = 1 To uniqueCount + 1 Step 1000
        batchRows = 1000
        If r + batchRows - 1 > uniqueCount + 1 Then batchRows = uniqueCount + 2 - r
        dst.Cells(r, 1).Resize(batchRows, 6 + tnos.Count).Value2 = work.Cells(r, 1).Resize(batchRows, 6 + tnos.Count).Value2
    Next r
    With dst.Range("A1").Resize(1, 6 + tnos.Count)
        .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True
    End With
    dst.Range("A1").Resize(uniqueCount + 1, 6 + tnos.Count).NumberFormat = "0"
    dst.Range("A1").Resize(1, 6 + tnos.Count).Columns.ColumnWidth = 12
    cfg.Range("D2").Value2 = path
    cfg.Range("D3").Value2 = Now
    cfg.Range("D4").Value2 = uniqueCount
    cfg.Range("D5").Value2 = accepted
    cfg.Range("D6").Value2 = skipped
    AddImportLog logs, runID, "INFO", "Complete", path, "", 0, 0, "Unique=" & uniqueCount & "; accepted=" & accepted & "; skipped=" & skipped & "; file bytes=" & fileSize, ""
    stage = "Write log"
    FlushImportLog logs
    MsgBox "Import complete. Unique addresses: " & uniqueCount & vbCrLf & "Skipped: " & skipped & vbCrLf & "See Import Log. Run ID: " & runID, vbInformation
Done:
    On Error Resume Next
    If fileOpen Then Close #fileNo
    If Not work Is Nothing Then
        Application.DisplayAlerts = False
        work.Delete
    End If
    Application.DisplayAlerts = oldAlerts
    Application.ScreenUpdating = oldScreen
    Application.StatusBar = oldStatus
    busy = False
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    On Error Resume Next
    Err.Clear
    AddImportLog logs, runID, "ERROR", stage, path, headerTNO, sourceLine, errorNumber, errorText, ""
    FlushImportLog logs
    If Err.Number <> 0 Then logError = vbCrLf & "Log could not be saved: " & Err.Description
    MsgBox "Import failed (" & errorNumber & "): " & errorText & vbCrLf & "Stage: " & stage & "; line: " & sourceLine & vbCrLf & "See Import Log. Run ID: " & runID & logError, vbExclamation
    GoTo Done
End Sub

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

Private Sub ParseASCBuffer(ByRef buffer As String, ByVal atEOF As Boolean, ByRef headerTNO As String, ByRef sourceLine As Long, ByRef selectedSections As Long, ByRef skipped As Long, ByVal tnos As Object, ByRef logs As Object, ByVal runID As String, ByVal path As String, ByRef records As Object)
    Dim re As Object, tokens As Object, token As Object, nums As Variant
    Dim s As String, reason As String, key As String, tno As String, tail As String
    Dim cursor As Long, tokenLine As Long, consumed As Long, openPos As Long
    Dim valid As Boolean
    records.RemoveAll
    If atEOF Then buffer = buffer & vbLf
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True: re.IgnoreCase = True
    re.Pattern = "\btno\s*=\s*[0-9]+(?=[^0-9])|\[[^\]]*\]"
    Set tokens = re.Execute(buffer)
    tokenLine = sourceLine
    For Each token In tokens
        s = Mid$(buffer, cursor + 1, token.FirstIndex - cursor)
        tokenLine = tokenLine + Len(s) - Len(Replace(s, vbLf, ""))
        s = CStr(token.Value)
        cursor = token.FirstIndex
        If Left$(s, 1) <> "[" Then
            headerTNO = CleanSpace(Mid$(s, InStr(s, "=") + 1))
            If tnos.Exists(headerTNO) Then selectedSections = selectedSections + 1
        ElseIf tnos.Exists(headerTNO) Then
            valid = TryParseAddress(s, nums, reason)
            If valid Then
                tno = CStr(nums(1))
                valid = (tno = headerTNO) And tnos.Exists(tno)
                If Not valid Then reason = "Bracket TNO differs from section TNO."
            End If
            If valid Then
                key = CStr(nums(2)) & "," & CStr(nums(3)) & "," & CStr(nums(4)) & "," & CStr(nums(5)) & "," & CStr(nums(6)) & "," & CStr(nums(7))
                records.Add records.Count, Array(key, tno, nums(2), nums(3), nums(4), nums(5), nums(6), nums(7))
            Else
                skipped = skipped + 1
                AddImportLog logs, runID, "WARN", "Parse addresses", path, headerTNO, tokenLine, 0, reason, s
            End If
        End If
        consumed = token.FirstIndex + token.Length
    Next token
    tail = Mid$(buffer, consumed + 1)
    openPos = InStr(tail, "[")
    If openPos > 0 Then
        consumed = consumed + openPos - 1
        tail = Mid$(buffer, consumed + 1)
        If atEOF Or Len(tail) > MAX_PENDING_RECORD Then
            If tnos.Exists(headerTNO) Then
                s = Left$(buffer, consumed)
                tokenLine = sourceLine + Len(s) - Len(Replace(s, vbLf, ""))
                reason = "Unterminated bracket at end of file."
                If Not atEOF Then reason = "Unterminated bracket exceeds 16384 characters; record discarded."
                AddImportLog logs, runID, "WARN", "Parse addresses", path, headerTNO, tokenLine, 0, reason, Left$(tail, 1024)
                skipped = skipped + 1
            End If
            consumed = Len(buffer)
        End If
    ElseIf atEOF Then
        consumed = Len(buffer)
    ElseIf Len(buffer) - consumed > 128 Then
        consumed = Len(buffer) - 128
    End If
    s = Left$(buffer, consumed)
    sourceLine = sourceLine + Len(s) - Len(Replace(s, vbLf, ""))
    buffer = Mid$(buffer, consumed + 1)
End Sub

' Commas define the fields. Whitespace around each complete integer is ignored.
Private Function TryParseAddress(ByVal bracket As String, ByRef nums As Variant, ByRef reason As String) As Boolean
    Dim fields As Variant, j As Long, n As Double
    reason = ""
    bracket = CleanSpace(bracket)
    If Left$(bracket, 1) <> "[" Or Right$(bracket, 1) <> "]" Then
        reason = "Missing brackets."
        Exit Function
    End If
    fields = Split(Mid$(bracket, 2, Len(bracket) - 2), ",")
    If UBound(fields) = 8 Then
        If CleanSpace(CStr(fields(8))) = "" Then ReDim Preserve fields(0 To 7)
    End If
    If UBound(fields) <> 7 Then
        reason = "Expected 8 comma-separated integer fields; found " & (UBound(fields) + 1) & "."
        Exit Function
    End If
    ReDim nums(0 To 7)
    For j = 0 To 7
        If Not IntegerInRange(CStr(fields(j)), -2147483648#, 2147483647#, n) Then
            reason = "Invalid integer at field " & (j + 1) & ": '" & Left$(CleanSpace(CStr(fields(j))), 80) & "'."
            Exit Function
        End If
        nums(j) = CLng(n)
    Next j
    If nums(1) < 100 Or nums(1) > 999 Then
        reason = "TNO=" & nums(1) & "; allowed 100-999.": Exit Function
    End If
    If nums(4) < 0 Or nums(4) > 255 Then
        reason = "BLK=" & nums(4) & "; allowed 0-255.": Exit Function
    End If
    If nums(5) < 0 Or nums(5) > 1023 Then
        reason = "WL=" & nums(5) & "; allowed 0-1023.": Exit Function
    End If
    If nums(6) < 0 Or nums(6) > 15 Then
        reason = "BL=" & nums(6) & "; allowed 0-15.": Exit Function
    End If
    If nums(7) < 0 Or nums(7) > 15 Then
        reason = "IO=" & nums(7) & "; allowed 0-15.": Exit Function
    End If
    TryParseAddress = True
End Function

' Logging helpers: append run history; do not overwrite previous diagnostics.
Private Sub AddImportLog(ByRef logs As Object, ByVal runID As String, ByVal level As String, ByVal stage As String, ByVal path As String, ByVal tno As String, ByVal sourceLine As Long, ByVal errorNumber As Long, ByVal reason As String, ByVal rawRecord As String)
    If logs Is Nothing Then Set logs = CreateObject("Scripting.Dictionary")
    logs.Add logs.Count, Array(Now, runID, level, stage, path, tno, sourceLine, errorNumber, Left$(reason, 32767), Left$(rawRecord, 32767))
End Sub

Private Sub FlushImportLog(ByRef logs As Object)
    Dim ws As Worksheet, row As Long, i As Long, j As Long, count As Long
    Dim data() As Variant, item As Variant
    If logs Is Nothing Then Exit Sub
    If logs.Count = 0 Then Exit Sub
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Import Log")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Import Log"
    End If
    count = logs.Count
    row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If row < 2 Then row = 2
    If row + count - 1 > ws.Rows.Count Then Err.Raise vbObjectError + 7, , "Import Log is full; archive older log rows before importing."
    ws.Range("A1:J1").Value2 = Array("Timestamp", "Run ID", "Level", "Stage", "ASC File", "TNO", "Source Line", "Error Code", "Reason", "Raw Record")
    ReDim data(1 To count, 1 To 10)
    For i = 0 To count - 1
        item = logs.Item(i)
        For j = 0 To 9: data(i + 1, j + 1) = item(j): Next j
    Next i
    ' Force text fields to text, including ASC snippets beginning with =,+,-,@.
    ws.Range("B" & row & ":F" & (row + count - 1)).NumberFormat = "@"
    ws.Range("I" & row & ":J" & (row + count - 1)).NumberFormat = "@"
    ws.Range("A" & row).Resize(count, 10).Value2 = data
    logs.RemoveAll
    ws.Range("A" & row).Resize(count, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Range("G" & row & ":H" & (row + count - 1)).NumberFormat = "0"
    With ws.Range("A1:J1")
        .Interior.Color = RGB(23, 50, 77)
        .Font.Color = vbWhite
        .Font.Bold = True
    End With
    ws.Columns("A:B").ColumnWidth = 24
    ws.Columns("C:D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 40
    ws.Columns("F:H").ColumnWidth = 14
    ws.Columns("I:J").ColumnWidth = 65
    ws.Range("I" & row & ":J" & (row + count - 1)).WrapText = True
    ws.Range("A" & row & ":J" & (row + count - 1)).EntireRow.AutoFit
End Sub

Private Function CleanSpace(ByVal text As String) As String
    text = Replace(Replace(Replace(text, vbTab, " "), vbCr, " "), vbLf, " ")
    text = Replace(Replace(Replace(text, ChrW(160), " "), ChrW(8239), " "), ChrW(12288), " ")
    text = Replace(text, ChrW(-257), "") ' UTF BOM U+FEFF
    CleanSpace = Trim$(text)
End Function

Private Function IntegerInRange(ByVal text As String, ByVal lower As Double, ByVal upper As Double, ByRef number As Double) As Boolean
    On Error GoTo Invalid
    text = CleanSpace(text)
    If mIntegerRegex Is Nothing Then
        Set mIntegerRegex = CreateObject("VBScript.RegExp")
        mIntegerRegex.Pattern = "^[+-]?[0-9]+$"
    End If
    If Not mIntegerRegex.Test(text) Then Exit Function
    number = CDbl(text)
    IntegerInRange = (number >= lower And number <= upper)
    Exit Function
Invalid:
    IntegerInRange = False
End Function
