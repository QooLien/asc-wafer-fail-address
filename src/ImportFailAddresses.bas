Attribute VB_Name = "ImportFailAddresses"
Option Explicit

Private Const READ_BYTES As Long = 262144
Private Const MAX_PENDING_RECORD As Long = 16384
Private Const MAX_RESULT_ROWS As Long = 1048575
Private mTokenRegex As Object
Private mBitMasks As Variant

' No logs. Keep one address index and one bit per selected TNO/address.
Public Sub ImportASCFailAddresses()
    Static busy As Boolean
    Dim cfg As Worksheet, dst As Worksheet, tnos As Object, addresses As Object
    Dim fd As Object, settings As Variant, headers() As Variant, data() As Variant
    Dim bits() As Byte, fields As Variant, v As Variant
    Dim raw As String, buffer As String, encoding As String, path As String
    Dim stage As String, headerTNO As String, s As String
    Dim r As Long, j As Long, b As Long, k As Long, rowIndex As Long, bitValue As Long
    Dim flagBytes As Long, capacity As Long, selectedSections As Long, columnCount As Long
    Dim skipped As Long, accepted As Long, uniqueCount As Long, chunkCount As Long
    Dim batchLimit As Long, batchSize As Long, batchRow As Long, written As Long
    Dim fileNo As Integer, fileOpen As Boolean, fileSize As Long, bytePosition As Long
    Dim n As Double, startedAt As Date, elapsedText As String, oldScreen As Boolean, oldEvents As Boolean
    Dim oldCalc As XlCalculation, oldStatus As Variant, errorNumber As Long, errorText As String
    If busy Then MsgBox "An import is already running.", vbExclamation: Exit Sub
    busy = True
    On Error GoTo Failed
    oldScreen = Application.ScreenUpdating
    oldEvents = Application.EnableEvents
    oldCalc = Application.Calculation
    oldStatus = Application.StatusBar
    mBitMasks = Array(1, 2, 4, 8, 16, 32, 64, 128)
    Set tnos = CreateObject("Scripting.Dictionary")
    Set addresses = CreateObject("Scripting.Dictionary")
    stage = "Validate Config"
    Set cfg = ThisWorkbook.Worksheets("Config")
    Set dst = ThisWorkbook.Worksheets("Fail Addresses")
    settings = cfg.Range("A2:A901").Value2
    For r = 1 To UBound(settings, 1)
        v = settings(r, 1)
        If IsError(v) Then Err.Raise vbObjectError + 1, , "Invalid TNO at Config!A" & (r + 1)
        s = CleanSpace(CStr(v))
        If Len(s) > 0 Then
            If Not IntegerInRange(s, 100, 999, n) Or Len(s) <> 3 Then Err.Raise vbObjectError + 2, , "TNO must be 100-999 at Config!A" & (r + 1)
            If Not tnos.Exists(s) Then tnos.Add s, tnos.Count
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
        If .Show <> -1 Then GoTo Done
        path = .SelectedItems(1)
    End With
    startedAt = Now
    stage = "Open ASC"
    If CreateObject("Scripting.FileSystemObject").GetFile(path).Size > 2147483647# Then Err.Raise vbObjectError + 8, , "ASC exceeds the 2 GB binary-reader limit. Split the input file."
    fileNo = FreeFile
    Open path For Binary Access Read As #fileNo
    fileOpen = True
    fileSize = LOF(fileNo)
    DetectEncoding fileNo, encoding
    flagBytes = (tnos.Count + 7) \ 8
    capacity = 4096
    ReDim bits(0 To flagBytes - 1, 0 To capacity - 1)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Do While Seek(fileNo) <= fileSize
        stage = "Read ASC chunk"
        raw = ReadASCChunk(fileNo, encoding)
        bytePosition = Seek(fileNo) - 1
        buffer = buffer & raw
        raw = ""
        stage = "Parse and merge addresses"
        ParseASCBuffer buffer, False, headerTNO, selectedSections, accepted, skipped, tnos, addresses, bits, flagBytes, capacity
        chunkCount = chunkCount + 1
        If chunkCount Mod 8 = 0 Then
            Application.StatusBar = "ASC: " & Format$(bytePosition / 1024, "#,##0") & " / " & Format$(fileSize / 1024, "#,##0") & " KB; unique=" & addresses.Count
            DoEvents
        End If
    Loop
    Close #fileNo
    fileOpen = False
    stage = "Parse final buffer"
    ParseASCBuffer buffer, True, headerTNO, selectedSections, accepted, skipped, tnos, addresses, bits, flagBytes, capacity
    If selectedSections = 0 Then Err.Raise vbObjectError + 5, , "No selected tno= section found. Check TNO and file encoding."
    If accepted = 0 And skipped > 0 Then Err.Raise vbObjectError + 6, , "No valid selected addresses. Previous results retained."
    uniqueCount = addresses.Count
    columnCount = 6 + tnos.Count
    ReDim headers(1 To 1, 1 To columnCount)
    fields = Array("X", "Y", "BLK", "WL", "BL", "IO")
    For j = 0 To 5: headers(1, j + 1) = fields(j): Next j
    For Each v In tnos: headers(1, 7 + CLng(tnos(v))) = CLng(v): Next v
    stage = "Write output batches"
    dst.UsedRange.ClearContents
    dst.Cells(1, 1).Resize(1, columnCount).Value2 = headers
    batchLimit = 131072 \ columnCount
    If batchLimit > 4096 Then batchLimit = 4096
    If uniqueCount > 0 Then
        batchSize = batchLimit
        If batchSize > uniqueCount Then batchSize = uniqueCount
        ReDim data(1 To batchSize, 1 To columnCount)
        For Each v In addresses
            rowIndex = CLng(addresses(v))
            batchRow = batchRow + 1
            fields = Split(CStr(v), ",")
            For j = 0 To 5: data(batchRow, j + 1) = CLng(fields(j)): Next j
            For b = 0 To flagBytes - 1
                bitValue = bits(b, rowIndex)
                If bitValue <> 0 Then
                    For k = 0 To 7
                        If (bitValue And mBitMasks(k)) <> 0 Then data(batchRow, 7 + b * 8 + k) = 1
                    Next k
                End If
            Next b
            If batchRow = batchSize Then
                dst.Cells(written + 2, 1).Resize(batchSize, columnCount).Value2 = data
                written = written + batchSize
                batchRow = 0
                If written < uniqueCount Then
                    batchSize = batchLimit
                    If batchSize > uniqueCount - written Then batchSize = uniqueCount - written
                    ReDim data(1 To batchSize, 1 To columnCount)
                End If
            End If
        Next v
    End If
    stage = "Format output"
    With dst.Cells(1, 1).Resize(uniqueCount + 1, columnCount)
        .NumberFormat = "0"
        .Font.Name = "Calibri": .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = False: .RowHeight = 26
    End With
    With dst.Cells(1, 1).Resize(1, columnCount)
        .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True
        .Columns.ColumnWidth = 12
    End With
    cfg.Range("D2").Value2 = path
    cfg.Range("D3").Value2 = Now
    cfg.Range("D4").Value2 = uniqueCount
    cfg.Range("D5").Value2 = accepted
    cfg.Range("D6").Value2 = skipped
    n = DateDiff("s", startedAt, Now)
    elapsedText = Format$(Fix(n / 3600), "00") & ":" & Format$(Fix(n / 60) Mod 60, "00") & ":" & Format$(n Mod 60, "00")
    MsgBox "Import complete. Unique addresses: " & uniqueCount & vbCrLf & "Run time (hh:mm:ss): " & elapsedText & vbCrLf & "Total seconds: " & Format$(n, "#,##0") & vbCrLf & "No log output.", vbInformation
Done:
    On Error Resume Next
    If fileOpen Then Close #fileNo
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.StatusBar = oldStatus
    busy = False
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    MsgBox "Import failed (" & errorNumber & "): " & errorText & vbCrLf & "Stage: " & stage & "; bytes read: " & bytePosition, vbExclamation
    Resume Done
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

Private Sub ParseASCBuffer(ByRef buffer As String, ByVal atEOF As Boolean, ByRef headerTNO As String, ByRef selectedSections As Long, ByRef accepted As Long, ByRef skipped As Long, ByVal tnos As Object, ByVal addresses As Object, ByRef bits() As Byte, ByVal flagBytes As Long, ByRef capacity As Long)
    Dim tokens As Object, token As Object, nums(0 To 7) As Long
    Dim s As String, key As String, tail As String, columnIndex As Long, rowIndex As Long
    Dim consumed As Long, openPos As Long, newCapacity As Long
    If atEOF Then buffer = buffer & vbLf
    If mTokenRegex Is Nothing Then
        Set mTokenRegex = CreateObject("VBScript.RegExp")
        mTokenRegex.Global = True: mTokenRegex.IgnoreCase = True
        mTokenRegex.Pattern = "\btno\s*=\s*[0-9]+(?=[^0-9])|\[[^\]]*\]"
    End If
    Set tokens = mTokenRegex.Execute(buffer)
    For Each token In tokens
        s = CStr(token.Value)
        If Left$(s, 1) <> "[" Then
            headerTNO = CleanSpace(Mid$(s, InStr(s, "=") + 1))
            If tnos.Exists(headerTNO) Then selectedSections = selectedSections + 1
        ElseIf tnos.Exists(headerTNO) Then
            If InStr(1, s, ",", vbBinaryCompare) = 0 Then
                skipped = skipped + 1
            ElseIf TryParseAddress(s, nums) Then
                If CStr(nums(1)) = headerTNO Then
                    key = CStr(nums(2)) & "," & CStr(nums(3)) & "," & CStr(nums(4)) & "," & CStr(nums(5)) & "," & CStr(nums(6)) & "," & CStr(nums(7))
                    If addresses.Exists(key) Then
                        rowIndex = CLng(addresses(key))
                    Else
                        rowIndex = addresses.Count
                        If rowIndex >= MAX_RESULT_ROWS Then Err.Raise vbObjectError + 4, , "Too many unique addresses for one Excel sheet."
                        If rowIndex >= capacity Then
                            newCapacity = capacity * 2
                            If newCapacity > MAX_RESULT_ROWS Then newCapacity = MAX_RESULT_ROWS
                            ReDim Preserve bits(0 To flagBytes - 1, 0 To newCapacity - 1)
                            capacity = newCapacity
                        End If
                        addresses.Add key, rowIndex
                    End If
                    columnIndex = CLng(tnos(headerTNO))
                    bits(columnIndex \ 8, rowIndex) = bits(columnIndex \ 8, rowIndex) Or mBitMasks(columnIndex And 7)
                    accepted = accepted + 1
                Else
                    skipped = skipped + 1
                End If
            Else
                skipped = skipped + 1
            End If
        End If
        consumed = token.FirstIndex + token.Length
    Next token
    tail = Mid$(buffer, consumed + 1)
    openPos = InStr(1, tail, "[", vbBinaryCompare)
    If openPos > 0 Then
        consumed = consumed + openPos - 1
        If atEOF Or Len(buffer) - consumed > MAX_PENDING_RECORD Then
            If tnos.Exists(headerTNO) Then skipped = skipped + 1
            consumed = Len(buffer)
        End If
    ElseIf atEOF Then
        consumed = Len(buffer)
    ElseIf Len(buffer) - consumed > 128 Then
        consumed = Len(buffer) - 128
    End If
    buffer = Mid$(buffer, consumed + 1)
End Sub

Private Function TryParseAddress(ByVal bracket As String, ByRef nums() As Long) As Boolean
    Dim fields As Variant, j As Long, count As Long, n As Double
    bracket = CleanSpace(bracket)
    If Left$(bracket, 1) <> "[" Or Right$(bracket, 1) <> "]" Then Exit Function
    fields = Split(Mid$(bracket, 2, Len(bracket) - 2), ",")
    count = UBound(fields)
    If count = 8 Then
        If Trim$(CStr(fields(8))) = "" Then count = 7
    End If
    If count <> 7 Then Exit Function
    For j = 0 To 7
        If Not IntegerInRange(CStr(fields(j)), -2147483648#, 2147483647#, n) Then Exit Function
        nums(j) = CLng(n)
    Next j
    If nums(1) < 100 Or nums(1) > 999 Then Exit Function
    If nums(4) < 0 Or nums(4) > 255 Then Exit Function
    If nums(5) < 0 Or nums(5) > 1023 Then Exit Function
    If nums(6) < 0 Or nums(6) > 15 Then Exit Function
    If nums(7) < 0 Or nums(7) > 15 Then Exit Function
    TryParseAddress = True
End Function

Private Function CleanSpace(ByVal text As String) As String
    If InStr(text, vbTab) > 0 Then text = Replace(text, vbTab, " ")
    If InStr(text, vbCr) > 0 Then text = Replace(text, vbCr, " ")
    If InStr(text, vbLf) > 0 Then text = Replace(text, vbLf, " ")
    If InStr(text, ChrW(160)) > 0 Then text = Replace(text, ChrW(160), " ")
    If InStr(text, ChrW(8239)) > 0 Then text = Replace(text, ChrW(8239), " ")
    If InStr(text, ChrW(12288)) > 0 Then text = Replace(text, ChrW(12288), " ")
    If InStr(text, ChrW(-257)) > 0 Then text = Replace(text, ChrW(-257), "")
    CleanSpace = Trim$(text)
End Function

' Validate a complete integer directly; avoid per-field regex and string conversion.
Private Function IntegerInRange(ByVal text As String, ByVal lower As Double, ByVal upper As Double, ByRef number As Double) As Boolean
    Dim j As Long, first As Long, digit As Long, negative As Boolean
    text = Trim$(text)
    If Len(text) = 0 Then Exit Function
    first = 1
    If Left$(text, 1) = "-" Then negative = True: first = 2
    If Left$(text, 1) = "+" Then first = 2
    If first > Len(text) Then Exit Function
    number = 0
    For j = first To Len(text)
        digit = AscW(Mid$(text, j, 1)) - 48
        If digit < 0 Or digit > 9 Then Exit Function
        number = number * 10 + digit
        If number > 2147483648# Then Exit Function
    Next j
    If negative Then number = -number
    IntegerInRange = (number >= lower And number <= upper)
End Function
