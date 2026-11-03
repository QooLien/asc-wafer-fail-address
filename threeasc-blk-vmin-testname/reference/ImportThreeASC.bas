Attribute VB_Name = "ImportFailAddresses"
Option Explicit

Private Const READ_BYTES As Long = 262144
Private Const MAX_PENDING_RECORD As Long = 16384
Private Const MAX_RESULT_ROWS As Long = 1048575
Private mTokenRegex As Object
Private mHeaderRegex As Object
Private mBitMasks As Variant

' No logs. Keep one address index and one bit per selected TNO/address.
Public Sub ImportASCFailAddresses()
    Static busy As Boolean
    Dim paths(1 To 3) As String, meta(1 To 3) As Object, steps(1 To 3) As Long
    Dim phase As Long, other As Long, swapPath As String, swapStep As Long, swapMeta As Object
    Dim localTNOs As Object, allAccepted As Long, allSkipped As Long, sourceStats(1 To 3, 1 To 6) As Variant
    Dim baseNames(1 To 3) As String, totalTNOs As Long, color As Long
    Dim cfg As Worksheet, dst As Worksheet, tnos As Object, addresses As Object
    Dim fd As Object, settings As Variant, headers() As Variant, data() As Variant
    Dim bits() As Byte, fields As Variant, v As Variant
    Dim raw As String, buffer As String, encoding As String, path As String
    Dim metadata As Object, metadataBuffer As String, outputPath As String, outputBase As String, fso As Object
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
    Set metadata = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
    stage = "Validate Config"
    Set cfg = ThisWorkbook.Worksheets("Config")
    Set dst = ThisWorkbook.Worksheets("Fail Address")
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
    stage = "Select three ASC files"
    Set fd = Application.FileDialog(3)
    With fd
        .Title = "Select exactly THREE ASC files": .AllowMultiSelect = True
        .Filters.Clear: .Filters.Add "ASC raw data", "*.asc"
        If .Show <> -1 Then GoTo Done
        If .SelectedItems.Count <> 3 Then Err.Raise vbObjectError + 20, , "Select exactly three ASC files."
        For phase = 1 To 3: paths(phase) = .SelectedItems(phase): Next phase
    End With
    startedAt = Now
    stage = "Validate LOT, WNO and STEP"
    For phase = 1 To 3
        For other = 1 To phase - 1
            If StrComp(paths(phase), paths(other), vbTextCompare) = 0 Then Err.Raise vbObjectError + 21, , "Duplicate ASC file."
        Next other
        Set meta(phase) = ReadFirstMetadata(paths(phase), fso)
        baseNames(phase) = BuildResultBaseName(meta(phase))
        steps(phase) = StepNumber(CStr(meta(phase)("STEP")))
        If phase > 1 Then
            If StrComp(CleanSpace(CStr(meta(phase)("LOT"))), CleanSpace(CStr(meta(1)("LOT"))), vbBinaryCompare) <> 0 Then Err.Raise vbObjectError + 22, , "Full LOT mismatch: " & paths(phase)
            If NormalWafer(CStr(meta(phase)("WNO"))) <> NormalWafer(CStr(meta(1)("WNO"))) Then Err.Raise vbObjectError + 23, , "WNO mismatch: " & paths(phase)
        End If
    Next phase
    For phase = 1 To 2
        For other = phase + 1 To 3
            If steps(phase) = steps(other) Then Err.Raise vbObjectError + 24, , "Duplicate numeric STEP; cannot assign WT1/WT2/WT3."
            If steps(phase) > steps(other) Then
                swapStep = steps(phase): steps(phase) = steps(other): steps(other) = swapStep
                swapPath = paths(phase): paths(phase) = paths(other): paths(other) = swapPath
                Set swapMeta = meta(phase): Set meta(phase) = meta(other): Set meta(other) = swapMeta
            End If
        Next other
    Next phase
    outputBase = SafeFilePart(Left$(CleanSpace(CStr(meta(1)("LOT"))), 6)) & "#" & NormalWafer(CStr(meta(1)("WNO"))) & "_WT1-WT3"
    outputPath = UniqueResultPath(fso, fso.GetParentFolderName(paths(1)), outputBase)
    totalTNOs = 3 * tnos.Count
    flagBytes = (totalTNOs + 7) \ 8
    capacity = 4096
    ReDim bits(0 To flagBytes - 1, 0 To capacity - 1)
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    For phase = 1 To 3
        path = paths(phase): buffer = "": headerTNO = ""
        selectedSections = 0: accepted = 0: skipped = 0: bytePosition = 0
        Set localTNOs = PhaseTNOs(tnos, phase)
        stage = "Open WT" & phase
        fileNo = FreeFile: Open path For Binary Access Read As #fileNo
        fileOpen = True: fileSize = LOF(fileNo)
        DetectEncoding fileNo, encoding
    Do While Seek(fileNo) <= fileSize
        stage = "Read ASC chunk"
        raw = ReadASCChunk(fileNo, encoding)
        bytePosition = Seek(fileNo) - 1
        buffer = buffer & raw
        raw = ""
        stage = "Parse and merge addresses"
        ParseASCBuffer buffer, False, headerTNO, selectedSections, accepted, skipped, localTNOs, addresses, bits, flagBytes, capacity
        chunkCount = chunkCount + 1
        If chunkCount Mod 8 = 0 Then
            Application.StatusBar = "ASC: " & Format$(bytePosition / 1024, "#,##0") & " / " & Format$(fileSize / 1024, "#,##0") & " KB; unique=" & addresses.Count
            DoEvents
        End If
    Loop
        Close #fileNo: fileOpen = False
        stage = "Parse final buffer WT" & phase
        ParseASCBuffer buffer, True, headerTNO, selectedSections, accepted, skipped, localTNOs, addresses, bits, flagBytes, capacity
        If selectedSections = 0 Then Err.Raise vbObjectError + 5, , "WT" & phase & ": no selected tno= section. File: " & path
        If accepted = 0 And skipped > 0 Then Err.Raise vbObjectError + 6, , "WT" & phase & ": no valid selected addresses. Previous results retained."
        sourceStats(phase, 1) = "WT" & phase
        sourceStats(phase, 2) = path
        sourceStats(phase, 3) = CStr(meta(phase)("STEP"))
        sourceStats(phase, 4) = BuildResultBaseName(meta(phase))
        sourceStats(phase, 5) = accepted: sourceStats(phase, 6) = skipped
        allAccepted = allAccepted + accepted: allSkipped = allSkipped + skipped
    Next phase
    uniqueCount = addresses.Count
    columnCount = 6 + totalTNOs
    ReDim headers(1 To 1, 1 To columnCount)
    fields = Array("X", "Y", "BLK", "WL", "BL", "IO")
    For j = 0 To 5: headers(1, j + 1) = fields(j): Next j
    For phase = 1 To 3
        Set localTNOs = PhaseTNOs(tnos, phase)
        For Each v In localTNOs: headers(1, 7 + CLng(localTNOs(v))) = "WT" & phase & vbLf & CStr(v): Next v
    Next phase
    stage = "Write output batches"
    If dst.AutoFilterMode Then dst.AutoFilterMode = False
    dst.UsedRange.FormatConditions.Delete
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
        .Orientation = 0: .RowHeight = 60
        .Columns.ColumnWidth = 10
    End With
    For phase = 1 To 3
        Select Case phase
            Case 1: color = RGB(219, 234, 254)
            Case 2: color = RGB(255, 237, 213)
            Case 3: color = RGB(220, 252, 231)
        End Select
        With dst.Cells(1, 7 + (phase - 1) * tnos.Count).Resize(1, tnos.Count)
            .Interior.Color = color: .Font.Color = RGB(23, 50, 77)
            .Orientation = 90: .WrapText = True: .Columns.ColumnWidth = 6
        End With
        If uniqueCount > 0 Then
            With dst.Cells(2, 7 + (phase - 1) * tnos.Count).Resize(uniqueCount, tnos.Count)
                .Interior.Pattern = xlNone
                With .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=1")
                    .Interior.Color = color
                End With
            End With
        End If
    Next phase
    cfg.Range("G2:L4").Value2 = sourceStats
    dst.Cells(1, 1).Resize(uniqueCount + 1, columnCount).AutoFilter
    cfg.Range("D2").Value2 = "3 ASC files; see WT source mapping"
    cfg.Range("D3").Value2 = Now
    cfg.Range("D4").Value2 = uniqueCount
    cfg.Range("D5").Value2 = allAccepted
    cfg.Range("D6").Value2 = allSkipped
    stage = "Save result workbook"
    Application.Calculation = oldCalc
    ThisWorkbook.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    n = DateDiff("s", startedAt, Now)
    elapsedText = Format$(Fix(n / 3600), "00") & ":" & Format$(Fix(n / 60) Mod 60, "00") & ":" & Format$(n Mod 60, "00")
    MsgBox "Import complete. Unique addresses: " & uniqueCount & vbCrLf & "Saved as: " & outputPath & vbCrLf & "Run time (hh:mm:ss): " & elapsedText & vbCrLf & "Total seconds: " & Format$(n, "#,##0") & vbCrLf & "No log output.", vbInformation
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


' Only orchestration helpers below are new. Original parser/reader helpers follow unchanged.
Private Function PhaseTNOs(ByVal selected As Object, ByVal phase As Long) As Object
    Dim result As Object, tno As Variant
    Set result = CreateObject("Scripting.Dictionary")
    For Each tno In selected
        result.Add CStr(tno), (phase - 1) * selected.Count + CLng(selected(tno))
    Next tno
    Set PhaseTNOs = result
End Function

Private Function StepNumber(ByVal value As String) As Long
    Dim re As Object, matches As Object, n As Double
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[^0-9]*([0-9]+)[^0-9]*$"
    Set matches = re.Execute(CleanSpace(value))
    If matches.Count <> 1 Then Err.Raise vbObjectError + 25, , "STEP must contain exactly one numeric sequence: " & value
    If Not IntegerInRange(CStr(matches(0).SubMatches(0)), 0, 2147483647#, n) Then Err.Raise vbObjectError + 25, , "STEP number out of range: " & value
    StepNumber = CLng(n)
End Function

Private Function NormalWafer(ByVal value As String) As String
    Dim n As Double
    value = CleanSpace(value)
    If Left$(value, 1) = "#" Then value = Mid$(value, 2)
    If Not IntegerInRange(value, 0, 2147483647#, n) Then Err.Raise vbObjectError + 11, , "WNO must be a non-negative integer."
    value = CStr(CLng(n))
    If Len(value) < 2 Then value = "0" & value
    NormalWafer = value
End Function

Private Function ReadFirstMetadata(ByVal path As String, ByVal fso As Object) As Object
    Dim no As Integer, opened As Boolean, encoding As String, text As String, result As Object
    Dim code As Long, description As String
    On Error GoTo Failed
    If fso.GetFile(path).Size > 2147483647# Then Err.Raise vbObjectError + 8, , "ASC exceeds the 2 GB binary-reader limit: " & path
    Set result = CreateObject("Scripting.Dictionary")
    no = FreeFile: Open path For Binary Access Read As #no: opened = True
    DetectEncoding no, encoding
    Do While Seek(no) <= LOF(no) And result.Count < 4
        text = text & ReadASCChunk(no, encoding)
        CaptureASCMetadata text, False, result
    Loop
    CaptureASCMetadata text, True, result
    Close #no: opened = False
    Set ReadFirstMetadata = result
    Exit Function
Failed:
    code = Err.Number: description = Err.Description
    On Error Resume Next
    If opened Then Close #no
    On Error GoTo 0
    Err.Raise code, , description
End Function

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

' First occurrence of each header wins, including an empty/invalid first value.
' The carry buffer is bounded; stop scanning once all four headers are captured.
Private Sub CaptureASCMetadata(ByRef text As String, ByVal atEOF As Boolean, ByVal metadata As Object)
    Dim matches As Object, item As Object, i As Long, startAt As Long, stopAt As Long
    Dim lineEnd As Long, crEnd As Long, keepAt As Long, field As String, value As String
    If metadata.Count = 4 Then text = "": Exit Sub
    If mHeaderRegex Is Nothing Then
        Set mHeaderRegex = CreateObject("VBScript.RegExp")
        mHeaderRegex.Global = True: mHeaderRegex.IgnoreCase = True
        mHeaderRegex.Pattern = "\b(LOT|WNO|STEP|Program)[ \t]*=[ \t]*"
    End If
    Set matches = mHeaderRegex.Execute(text)
    For i = 0 To matches.Count - 1
        Set item = matches(i)
        field = UCase$(item.SubMatches(0))
        If Not metadata.Exists(field) Then
            startAt = item.FirstIndex + item.Length + 1
            lineEnd = InStr(startAt, text, vbLf, vbBinaryCompare)
            crEnd = InStr(startAt, text, vbCr, vbBinaryCompare)
            If lineEnd = 0 Or (crEnd > 0 And crEnd < lineEnd) Then lineEnd = crEnd
            stopAt = lineEnd
            If i < matches.Count - 1 Then
                If stopAt = 0 Or matches(i + 1).FirstIndex + 1 < stopAt Then stopAt = matches(i + 1).FirstIndex + 1
            End If
            If stopAt = 0 And Not atEOF Then
                keepAt = item.FirstIndex + 1
                If Len(text) - item.FirstIndex > MAX_PENDING_RECORD Then
                    metadata.Add field, ""
                    keepAt = 0
                End If
                Exit For
            End If
            If stopAt = 0 Then stopAt = Len(text) + 1
            value = CleanSpace(Mid$(text, startAt, stopAt - startAt))
            If field <> "PROGRAM" Then value = FirstHeaderToken(value)
            metadata.Add field, value
            If metadata.Count = 4 Then Exit For
        End If
    Next i
    If metadata.Count = 4 Or atEOF Then
        text = ""
    ElseIf keepAt > 0 Then
        text = Mid$(text, keepAt)
    ElseIf Len(text) > 128 Then
        text = Right$(text, 128)
    End If
End Sub

Private Function FirstHeaderToken(ByVal value As String) As String
    Dim quote As String, p As Long, j As Long, c As String
    value = CleanSpace(value)
    If Len(value) = 0 Then Exit Function
    quote = Left$(value, 1)
    If quote = Chr$(34) Or quote = "'" Then
        p = InStr(2, value, quote, vbBinaryCompare)
        If p > 0 Then FirstHeaderToken = Mid$(value, 2, p - 2)
        Exit Function
    End If
    For j = 1 To Len(value)
        c = Mid$(value, j, 1)
        If c = " " Or c = "," Or c = ";" Then Exit For
    Next j
    FirstHeaderToken = Left$(value, j - 1)
End Function

Private Function BuildResultBaseName(ByVal metadata As Object) As String
    Dim field As Variant, lot As String, wafer As String, stepValue As String
    Dim program As String, version As String, versionPos As Long, j As Long, letter As Long, n As Double
    For Each field In Array("LOT", "WNO", "STEP", "PROGRAM")
        If Not metadata.Exists(field) Then Err.Raise vbObjectError + 9, , "Missing first ASC header: " & field & ". Results were not replaced."
        If Len(CleanSpace(CStr(metadata(field)))) = 0 Then Err.Raise vbObjectError + 10, , "First ASC header is empty or too long: " & field
    Next field
    lot = SafeFilePart(Left$(CleanSpace(CStr(metadata("LOT"))), 6))
    wafer = CleanSpace(CStr(metadata("WNO")))
    If Left$(wafer, 1) = "#" Then wafer = Mid$(wafer, 2)
    If Not IntegerInRange(wafer, 0, 2147483647#, n) Then Err.Raise vbObjectError + 11, , "First WNO must be a non-negative integer."
    wafer = CStr(CLng(n))
    If Len(wafer) < 2 Then wafer = "0" & wafer
    stepValue = CleanSpace(CStr(metadata("STEP")))
    If UCase$(Left$(stepValue, 3)) = "WTT" Then stepValue = Mid$(stepValue, 4)
    stepValue = SafeFilePart(stepValue)
    program = CStr(metadata("PROGRAM"))
    versionPos = InStr(1, program, "VER_", vbTextCompare)
    If versionPos = 0 Then Err.Raise vbObjectError + 12, , "No VER_ version found in the first Program header."
    version = UCase$(Mid$(program, versionPos + 4, 2))
    If Len(version) <> 2 Then Err.Raise vbObjectError + 12, , "First VER_ must be followed by two English letters."
    For j = 1 To 2
        letter = AscW(Mid$(version, j, 1))
        If letter < 65 Or letter > 90 Then Err.Raise vbObjectError + 12, , "First VER_ must be followed by two English letters."
    Next j
    program = "Ver_" & version
    BuildResultBaseName = lot & "#" & wafer & "_" & stepValue & "_" & program
End Function

Private Function SafeFilePart(ByVal value As String) As String
    Dim bad As Variant, j As Long
    For Each bad In Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|", "[", "]")
        value = Replace(value, CStr(bad), "_")
    Next bad
    For j = 0 To 31: value = Replace(value, Chr$(j), "_"): Next j
    value = Trim$(value)
    Do While Len(value) > 0
        If Right$(value, 1) <> "." And Right$(value, 1) <> " " Then Exit Do
        value = Left$(value, Len(value) - 1)
    Loop
    If Len(value) = 0 Then Err.Raise vbObjectError + 13, , "Empty filename component after normalization."
    SafeFilePart = value
End Function

Private Function UniqueResultPath(ByVal fso As Object, ByVal folder As String, ByVal base As String) As String
    Dim candidate As String, sequence As Long, suffix As String
    candidate = fso.BuildPath(folder, base & ".xlsm")
    Do While fso.FileExists(candidate)
        sequence = sequence + 1
        If sequence > 9999 Then Err.Raise vbObjectError + 14, , "Too many files with this result name."
        suffix = CStr(sequence)
        If sequence < 1000 Then suffix = Right$("000" & suffix, 3)
        candidate = fso.BuildPath(folder, base & "_" & suffix & ".xlsm")
    Loop
    If Len(candidate) > 218 Then Err.Raise vbObjectError + 15, , "Result path exceeds 218 characters. Move the ASC to a shorter folder path."
    UniqueResultPath = candidate
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
