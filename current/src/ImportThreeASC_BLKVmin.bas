Attribute VB_Name = "ImportFailAddresses"
Option Explicit

Private Const READ_BYTES As Long = 262144
Private Const MAX_PENDING_RECORD As Long = 16384
Private Const MAX_RESULT_ROWS As Long = 1048575
Private mTokenRegex As Object
Private mHeaderRegex As Object
Private mBitMasks As Variant
Private busy As Boolean

' No logs. Keep one address index and one bit per selected TNO/address.
Public Sub ImportASCFailAddresses()
    RunThreeASC False
End Sub

Public Sub ImportASCFBC()
    RunThreeASC True
End Sub

Public Sub BuildFailMaps()
    Dim text As String, started As Date, oldScreen As Boolean, oldEvents As Boolean
    If busy Then MsgBox "An import or Map build is already running.", vbExclamation: Exit Sub
    busy = True: started = Now
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    On Error GoTo FailedMap
    Application.ScreenUpdating = False: Application.EnableEvents = False
    text = BLKVminEngine.BLK_BuildFailMaps(ThisWorkbook)
    BLKVminEngine.BLK_ApplyPresentation ThisWorkbook
    ThisWorkbook.Worksheets("Config").Range("T2").Value2 = text
    If Len(text) > 0 Then
        MsgBox text, vbExclamation
    Else
        ThisWorkbook.Save
        MsgBox "Fail Maps complete. Seconds: " & DateDiff("s", started, Now), vbInformation
    End If
DoneMap:
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents: busy = False
    Exit Sub
FailedMap:
    MsgBox "Map failed (" & Err.Number & "): " & Err.Description, vbExclamation
    Resume DoneMap
End Sub

Private Sub RunThreeASC(ByVal fbcOnly As Boolean)
    Dim paths(1 To 3) As String, meta(1 To 3) As Object, steps(1 To 3) As Long
    Dim phase As Long, other As Long, swapPath As String, swapStep As Long, swapMeta As Object
    Dim localTNOs As Object, allAccepted As Long, allSkipped As Long, sourceStats(1 To 3, 1 To 6) As Variant
    Dim baseNames(1 To 3) As String, totalTNOs As Long, color As Long
    Dim blkWarnings As String, blkMessage As String, addressLabels As Object
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
    If Not fbcOnly Then Set dst = ThisWorkbook.Worksheets("Fail Address")
    stage = "Validate Monitor / Function settings"
    BLKVminEngine.BLK_Setup cfg
    Set tnos = BLKVminEngine.BLK_AddressTNOs()
    If Not fbcOnly Then Set addressLabels = BuildAddressLabels(cfg)
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
    outputBase = SafeFilePart(Left$(CleanSpace(CStr(meta(1)("LOT"))), 6)) & "#" & NormalWafer(CStr(meta(1)("WNO"))) & "_WT1-WT3_FBC"
    If fbcOnly Then
        outputBase = outputBase & "_ONLY"
    Else
        outputBase = Replace(outputBase, "_WT1-WT3_FBC", "_WT1-WT3_Address")
    End If
    outputPath = UniqueResultPath(fso, fso.GetParentFolderName(paths(1)), outputBase)
    totalTNOs = 3 * tnos.Count
    If Not fbcOnly Then
        flagBytes = (totalTNOs + 7) \ 8
        capacity = 4096
        ReDim bits(0 To flagBytes - 1, 0 To capacity - 1)
    End If
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    ThisWorkbook.Worksheets("Import Index").Cells(IIf(fbcOnly, 5, 2), 8).Value2 = IIf(fbcOnly, "FBC Pending", "Address Pending")
    For phase = 1 To 3
        path = paths(phase): buffer = "": headerTNO = ""
        selectedSections = 0: accepted = 0: skipped = 0: bytePosition = 0
        Set localTNOs = PhaseTNOs(tnos, phase)
        stage = "Open WT" & phase
        fileNo = FreeFile: Open path For Binary Access Read As #fileNo
        fileOpen = True: fileSize = LOF(fileNo)
        DetectEncoding fileNo, encoding
        If fbcOnly Then BLKVminEngine.BLK_BeginFile path
    Do While Seek(fileNo) <= fileSize
        stage = "Read ASC chunk"
        raw = ReadASCChunk(fileNo, encoding)
        bytePosition = Seek(fileNo) - 1
        If fbcOnly Then
            BLKVminEngine.BLK_ObserveLocations raw, False
            BLKVminEngine.BLK_Feed raw, False
        Else
            buffer = buffer & raw
        End If
        raw = ""
        If fbcOnly Then stage = "FBC stream WT" & phase Else stage = "Parse and merge addresses"
        If Not fbcOnly Then ParseASCBuffer buffer, False, headerTNO, selectedSections, accepted, skipped, localTNOs, addresses, bits, flagBytes, capacity
        chunkCount = chunkCount + 1
        If chunkCount Mod 8 = 0 Then
            Application.StatusBar = "ASC: " & Format$(bytePosition / 1024, "#,##0") & " / " & Format$(fileSize / 1024, "#,##0") & " KB"
            DoEvents
        End If
    Loop
        Close #fileNo: fileOpen = False
        stage = "Parse final buffer WT" & phase
        If Not fbcOnly Then ParseASCBuffer buffer, True, headerTNO, selectedSections, accepted, skipped, localTNOs, addresses, bits, flagBytes, capacity
        If Not fbcOnly And selectedSections = 0 Then Err.Raise vbObjectError + 5, , "WT" & phase & ": no selected tno= section. File: " & path
        If Not fbcOnly And accepted = 0 And skipped > 0 Then Err.Raise vbObjectError + 6, , "WT" & phase & ": no valid selected addresses. Previous results retained."
        sourceStats(phase, 1) = "WT" & phase
        sourceStats(phase, 2) = path
        sourceStats(phase, 3) = CStr(meta(phase)("STEP"))
        sourceStats(phase, 4) = BuildResultBaseName(meta(phase))
        sourceStats(phase, 5) = accepted: sourceStats(phase, 6) = skipped
        If fbcOnly Then
            stage = "FBC by BLK WT" & phase
            BLKVminEngine.BLK_ObserveLocations "", True
            BLKVminEngine.BLK_Feed "", True
            accepted = BLKVminEngine.BLK_RecordCount()
            sourceStats(phase, 5) = accepted
            blkMessage = BLKVminEngine.BLK_WritePhase(ThisWorkbook, phase)
            If Len(blkMessage) > 0 Then blkWarnings = blkWarnings & vbCrLf & "WT" & phase & ": " & blkMessage
        End If
        allAccepted = allAccepted + accepted: allSkipped = allSkipped + skipped
    Next phase
    If Not fbcOnly Then
    uniqueCount = addresses.Count
    columnCount = 6 + totalTNOs
    ReDim headers(1 To 1, 1 To columnCount)
    fields = Array("X", "Y", "BLK", "WL", "BL", "IO")
    For j = 0 To 5: headers(1, j + 1) = fields(j): Next j
    For phase = 1 To 3
        Set localTNOs = PhaseTNOs(tnos, phase)
        For Each v In localTNOs: headers(1, 7 + CLng(localTNOs(v))) = "WT" & phase & vbLf & CStr(addressLabels(CStr(v))): Next v
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
        .Orientation = 0: .RowHeight = 150
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
    cfg.Range("J2:O4").Value2 = sourceStats
    dst.Cells(1, 1).Resize(uniqueCount + 1, columnCount).AutoFilter
    cfg.Range("I2").Value2 = "3 ASC files; see WT source mapping"
    cfg.Range("I3").Value2 = Now
    cfg.Range("I4").Value2 = uniqueCount
    cfg.Range("I5").Value2 = allAccepted
    cfg.Range("I6").Value2 = allSkipped
    Else
        cfg.Range("J2:O4").Value2 = sourceStats
        cfg.Range("I2").Value2 = "FBC only: 3 ASC files"
        cfg.Range("I3").Value2 = Now
        blkMessage = BLKVminEngine.BLK_Finalize(ThisWorkbook)
        If Len(blkMessage) > 0 Then
            blkWarnings = blkWarnings & vbCrLf & blkMessage & "; Maps were not updated."
        Else
            blkMessage = BLKVminEngine.BLK_BuildFailMaps(ThisWorkbook)
            If Len(blkMessage) > 0 Then blkWarnings = blkWarnings & vbCrLf & blkMessage
        End If
        cfg.Range("T2").Value2 = blkWarnings
    End If
    RecordImportIndex ThisWorkbook, IIf(fbcOnly, "FBC", "Address"), meta, paths, tnos
    BLKVminEngine.BLK_ApplyPresentation ThisWorkbook
    stage = "Save result workbook"
    Application.Calculation = oldCalc
    ThisWorkbook.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    n = DateDiff("s", startedAt, Now)
    elapsedText = Format$(Fix(n / 3600), "00") & ":" & Format$(Fix(n / 60) Mod 60, "00") & ":" & Format$(n Mod 60, "00")
    If fbcOnly Then
        MsgBox "FBC import complete. Records: " & allAccepted & vbCrLf & "Saved as: " & outputPath & vbCrLf & "Run time (hh:mm:ss): " & elapsedText & vbCrLf & "Fail Address was not changed." & blkWarnings, vbInformation
    Else
    MsgBox "Import complete. Unique addresses: " & uniqueCount & vbCrLf & "Saved as: " & outputPath & vbCrLf & "Run time (hh:mm:ss): " & elapsedText & vbCrLf & "Total seconds: " & Format$(n, "#,##0") & vbCrLf & "No log output." & blkWarnings, vbInformation
    End If
Done:
    On Error Resume Next
    If fileOpen Then Close #fileNo
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    Application.StatusBar = oldStatus
    BLKVminEngine.BLK_Release
    busy = False
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    MsgBox "Import failed (" & errorNumber & "): " & errorText & vbCrLf & "Stage: " & stage & "; bytes read: " & bytePosition, vbExclamation
    Resume Done
End Sub


Private Function BuildAddressLabels(ByVal cfg As Worksheet) As Object
    Dim labels As Object
    Set labels = CreateObject("Scripting.Dictionary")
    AddAddressLabels cfg.Range("A2:C901").Value2, labels
    AddAddressLabels cfg.Range("D2:F901").Value2, labels
    Set BuildAddressLabels = labels
End Function

Private Sub AddAddressLabels(ByVal data As Variant, ByVal labels As Object)
    Dim i As Long, tno As String, name As String, label As String
    For i = LBound(data, 1) To UBound(data, 1)
        If Len(Trim$(CStr(data(i, 1)))) > 0 Then
            tno = CStr(CLng(data(i, 1))): label = tno
            If Not IsError(data(i, 2)) And Not IsError(data(i, 3)) Then
                name = Trim$(CStr(data(i, 2)))
                If Len(name) > 0 And Len(Trim$(CStr(data(i, 3)))) > 0 And IsNumeric(data(i, 3)) Then
                    label = name & "_" & AddressVoltageLabel(CDbl(data(i, 3)))
                End If
            End If
            labels.Add tno, label
        End If
    Next i
End Sub

Private Function AddressVoltageLabel(ByVal voltage As Double) As String
    Dim value As Double, whole As Double, fraction As String, prefix As String
    value = Round(Abs(voltage), 9): whole = Fix(value)
    fraction = Right$("000000000" & CStr(CLng(Round((value - whole) * 1000000000#, 0))), 9)
    Do While Len(fraction) > 2
        If Right$(fraction, 1) <> "0" Then Exit Do
        fraction = Left$(fraction, Len(fraction) - 1)
    Loop
    If voltage < 0 Then prefix = "-"
    AddressVoltageLabel = prefix & CStr(whole) & "." & fraction & "v"
End Function

Public Sub FindMissingBLKAddresses()
    Dim result As String, oldScreen As Boolean, oldEvents As Boolean, oldCalc As XlCalculation, oldStatus As Variant, started As Date
    Dim source As Workbook, candidate As Workbook, fd As Object, path As String, opened As Boolean, oldSecurity As Long
    If busy Then MsgBox "An import, Map or audit is already running.", vbExclamation: Exit Sub
    busy = True: started = Now
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents: oldCalc = Application.Calculation: oldStatus = Application.StatusBar: oldSecurity = Application.AutomationSecurity
    On Error GoTo FailedAudit
    Application.ScreenUpdating = False: Application.EnableEvents = False: Application.Calculation = xlCalculationManual
    Set source = ThisWorkbook
    If AuditLastRow(source.Worksheets("Fail Address"), 1) < 2 Or AuditLastRow(source.Worksheets("FBC by BLK"), 2) < 2 Then
        Set fd = Application.FileDialog(3)
        With fd
            .Title = "Select an existing Excel result containing Config, Fail Address and FBC by BLK"
            .AllowMultiSelect = False: .Filters.Clear: .Filters.Add "Excel results", "*.xlsm;*.xlsx"
            If .Show <> -1 Then GoTo DoneAudit
            path = .SelectedItems(1)
        End With
        Set source = Nothing
        For Each candidate In Application.Workbooks
            If StrComp(candidate.FullName, path, vbTextCompare) = 0 Then Set source = candidate: Exit For
        Next candidate
        If source Is Nothing Then
            Application.AutomationSecurity = 3
            Set source = Application.Workbooks.Open(Filename:=path, UpdateLinks:=0, ReadOnly:=True)
            Application.AutomationSecurity = oldSecurity: opened = True
        End If
    End If
    result = AuditFunctionBlocks(source, ThisWorkbook)
    If Left$(result, 12) = "Audit failed" Then
        MsgBox result, vbExclamation
    Else
        ThisWorkbook.Save
        MsgBox result & vbCrLf & "Seconds: " & DateDiff("s", started, Now), vbInformation
    End If
DoneAudit:
    On Error Resume Next
    If opened Then source.Close SaveChanges:=False
    Application.AutomationSecurity = oldSecurity
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents: Application.Calculation = oldCalc: Application.StatusBar = oldStatus
    busy = False: Exit Sub
FailedAudit:
    MsgBox "Audit failed (" & Err.Number & "): " & Err.Description, vbExclamation
    Resume DoneAudit
End Sub

Private Function AuditFunctions(ByVal cfg As Worksheet, Optional ByVal scopeOverride As String = "") As Object
    Dim result As Object, data As Variant, i As Long, tno As String, stopOnly As Boolean, scope As String, func As String, voltage As Variant, pattern As String
    Set result = CreateObject("Scripting.Dictionary")
    If Len(scopeOverride) > 0 Then scope = UCase$(Trim$(scopeOverride)) Else scope = UCase$(Trim$(CStr(cfg.Range("Q2").Value2)))
    If scope <> "" And scope <> "ALL FUNCTION" And scope <> "STOP SIGN" Then Err.Raise vbObjectError + 480, , "Audit scope Q2 must be All Function or Stop Sign."
    stopOnly = (scope = "STOP SIGN")
    data = cfg.Range("D2:G901").Value2
    For i = LBound(data, 1) To UBound(data, 1)
        If IsError(data(i, 1)) Or IsError(data(i, 4)) Then Err.Raise vbObjectError + 480, , "Function configuration error at row " & i + 1
        tno = Trim$(CStr(data(i, 1))): pattern = Trim$(CStr(data(i, 4)))
        If Len(tno) > 0 And (Not stopOnly Or Len(pattern) > 0) Then
            If Len(tno) <> 3 Then Err.Raise vbObjectError + 480, , "Function TNO must be100-999."
            tno = CStr(AuditInteger(tno, 100, 999))
            If result.Exists(tno) Then Err.Raise vbObjectError + 480, , "Duplicate Function TNO: " & tno
            func = "": voltage = Empty
            If Not IsError(data(i, 2)) Then func = CStr(data(i, 2))
            If Not IsError(data(i, 3)) Then voltage = data(i, 3)
            result.Add tno, Array(func, voltage, pattern)
        End If
    Next i
    If result.Count = 0 Then Err.Raise vbObjectError + 480, , "No Function TNOs in audit scope. Enter TNOs in Config D."
    Set AuditFunctions = result
End Function

Private Function AuditInteger(ByVal value As Variant, ByVal minimum As Double, ByVal maximum As Double) As Long
    If IsError(value) Then Err.Raise vbObjectError + 481, , "Cell error in comparison input."
    If Len(Trim$(CStr(value))) = 0 Or Not IsNumeric(value) Then Err.Raise vbObjectError + 481, , "Missing/non-numeric comparison coordinate or TNO."
    If CDbl(value) <> Fix(CDbl(value)) Or CDbl(value) < minimum Or CDbl(value) > maximum Then Err.Raise vbObjectError + 481, , "Comparison integer out of range: " & CStr(value)
    AuditInteger = CLng(value)
End Function

Private Function AuditLastRow(ByVal ws As Worksheet, ByVal column As Long) As Long
    Dim last As Long, first As Long, take As Long, i As Long, data As Variant, value As Variant
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    Do While last >= 1
        first = last - 8191: If first < 1 Then first = 1
        take = last - first + 1
        data = ws.Cells(first, column).Resize(take, 1).Value2
        For i = take To 1 Step -1
            If take = 1 Then value = data Else value = data(i, 1)
            If IsError(value) Then AuditLastRow = first + i - 1: Exit Function
            If Len(CStr(value)) > 0 Then AuditLastRow = first + i - 1: Exit Function
        Next i
        last = first - 1
    Loop
End Function

Private Function AuditFBCValue(ByVal value As Variant) As Double
    If IsError(value) Then Err.Raise vbObjectError + 481, , "FBC cell error."
    If Len(Trim$(CStr(value))) = 0 Or Not IsNumeric(value) Then Err.Raise vbObjectError + 481, , "Missing/non-numeric FBC."
    If CDbl(value) <> Fix(CDbl(value)) Or CDbl(value) < 0 Or CDbl(value) > 9007199254740991# Then Err.Raise vbObjectError + 481, , "Invalid FBC integer."
    AuditFBCValue = CDbl(value)
End Function

Private Function AuditHeaderWT(ByVal text As String, ByRef label As String) As String
    Dim p As Long, wt As String
    text = Replace(text, vbCr, ""): p = InStr(text, vbLf)
    If p = 0 Then Exit Function
    wt = UCase$(Trim$(Left$(text, p - 1)))
    If wt <> "WT1" And wt <> "WT2" And wt <> "WT3" Then Exit Function
    label = Trim$(Mid$(text, p + 1)): AuditHeaderWT = wt
End Function

Private Function ResolveAuditColumns(ByVal wb As Workbook, ByVal wanted As Object, ByRef note As String) As Object
    Dim result As Object, coverage As Object, labels As Object, labelCounts As Object, index As Worksheet, ws As Worksheet
    Dim headers As Variant, data As Variant, last As Long, i As Long, col As Long, phase As Long, wt As String, tno As Variant, label As String, key As String, v As Variant
    Set result = CreateObject("Scripting.Dictionary"): Set coverage = CreateObject("Scripting.Dictionary")
    Set ws = wb.Worksheets("Fail Address")
    last = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If last < 7 Then Err.Raise vbObjectError + 482, , "Fail Address has no imported TNO columns."
    headers = ws.Cells(1, 1).Resize(1, last).Value2
    On Error Resume Next
    Set index = wb.Worksheets("Import Index")
    On Error GoTo 0
    If Not index Is Nothing Then
        i = index.Cells(index.Rows.Count, 1).End(xlUp).Row
        If i >= 2 Then
            data = index.Range("A2:D" & i).Value2
            For i = 1 To UBound(data, 1)
                If wanted.Exists(CStr(data(i, 2))) Then
                    wt = CStr(data(i, 1)): col = AuditInteger(data(i, 3), 7, last)
                    If wt <> "WT1" And wt <> "WT2" And wt <> "WT3" Then Err.Raise vbObjectError + 482, , "Invalid WT in numeric Address index."
                    If CStr(headers(1, col)) <> CStr(data(i, 4)) Then Err.Raise vbObjectError + 482, , "Address header changed since numeric TNO index was recorded; reimport Address."
                    key = wt & "|" & CStr(data(i, 2))
                    If coverage.Exists(key) Or result.Exists(CStr(col)) Then Err.Raise vbObjectError + 482, , "Duplicate Address TNO/column mapping."
                    coverage.Add key, True: result.Add CStr(col), Array(wt, CStr(data(i, 2)))
                End If
            Next i
            note = "Numeric TNO import index verified."
        End If
    End If
    If Len(note) = 0 Then
        Set labels = BuildAddressLabels(wb.Worksheets("Config"))
        Set labelCounts = CreateObject("Scripting.Dictionary"): labelCounts.CompareMode = vbTextCompare
        For Each v In labels
            label = CStr(labels(v))
            If labelCounts.Exists(label) Then labelCounts(label) = CLng(labelCounts(label)) + 1 Else labelCounts.Add label, 1
        Next v
        For col = 7 To last
            If IsError(headers(1, col)) Then Err.Raise vbObjectError + 482, , "Address header cell error."
            wt = AuditHeaderWT(CStr(headers(1, col)), label)
            If Len(wt) > 0 Then
                For Each tno In wanted
                    If label = CStr(tno) Or StrComp(label, CStr(labels(tno)), vbTextCompare) = 0 Then
                        If label <> CStr(tno) And CLng(labelCounts(CStr(labels(tno)))) > 1 Then Err.Raise vbObjectError + 482, , "Legacy Address has ambiguous Func/Voltage headers. Import Address once with this version to store numeric TNO mapping."
                        key = wt & "|" & CStr(tno)
                        If coverage.Exists(key) Or result.Exists(CStr(col)) Then Err.Raise vbObjectError + 482, , "Ambiguous Address column mapping."
                        coverage.Add key, True: result.Add CStr(col), Array(wt, CStr(tno))
                    End If
                Next tno
            End If
        Next col
        note = "Legacy Address: unique header labels resolved against Config."
    End If
    For Each tno In wanted
        For phase = 1 To 3
            If Not coverage.Exists("WT" & phase & "|" & tno) Then Err.Raise vbObjectError + 482, , "Address has no verified column for WT" & phase & "/TNO=" & tno
        Next phase
    Next tno
    Set ResolveAuditColumns = result
End Function

Private Function AuditSourceNote(ByVal wb As Workbook) As String
    Dim index As Worksheet, r As Long, c As Long, a As Long, f As Long, identity As Boolean
    On Error Resume Next
    Set index = wb.Worksheets("Import Index")
    On Error GoTo 0
    If index Is Nothing Then AuditSourceNote = "Legacy source identity not recorded; verify both outputs belong to the same Raw files.": Exit Function
    For r = 2 To 7
        If InStr(1, CStr(index.Cells(r, 8).Value2), "Pending", vbTextCompare) > 0 Then Err.Raise vbObjectError + 483, , "An Address/FBC import is incomplete; finish the import before comparison."
    Next r
    identity = (CStr(index.Range("H2").Value2) = "Address" And CStr(index.Range("H5").Value2) = "FBC")
    If Not identity Then AuditSourceNote = "Legacy source identity not recorded; verify both outputs belong to the same Raw files.": Exit Function
    For r = 0 To 2
        a = 2 + r: f = 5 + r
        For c = 9 To 13
            If CStr(index.Cells(a, c).Value2) <> CStr(index.Cells(f, c).Value2) Then Err.Raise vbObjectError + 483, , "Address/FBC source LOT/WNO/STEP/version mismatch for WT" & r + 1
        Next c
    Next r
    AuditSourceNote = "Address/FBC source LOT, WNO, STEP and identity verified."
End Function

Private Sub BuildDetailedBLKIndex(ByVal ws As Worksheet, ByVal columns As Object, ByVal blocks As Object, ByVal totals As Object)
    Dim groups As New Collection, lastCol As Long, lastRow As Long, col As Long, first As Long, width As Long, maxWidth As Long
    Dim g As Variant, info As Variant, coords As Variant, flags As Variant, suffix() As String, chips() As String
    Dim start As Long, take As Long, batch As Long, r As Long, j As Long, flag As Variant, key As String, chip As String, x As Long, y As Long, b As Long
    Dim blockNos() As Long, bits As Variant, activeChip As String, activeTotal As Long, emptyBits(0 To 31) As Byte
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    col = 7
    Do While col <= lastCol
        If columns.Exists(CStr(col)) Then
            first = col
            Do While col <= lastCol
                If Not columns.Exists(CStr(col)) Then Exit Do
                col = col + 1
            Loop
            width = col - first: groups.Add Array(first, width)
            If width > maxWidth Then maxWidth = width
        Else
            col = col + 1
        End If
    Loop
    batch = 4096: If batch * maxWidth > 131072 Then batch = 131072 \ maxWidth
    lastRow = AuditLastRow(ws, 1)
    For start = 2 To lastRow Step batch
        take = lastRow - start + 1: If take > batch Then take = batch
        coords = ws.Cells(start, 1).Resize(take, 3).Value2
        ReDim suffix(1 To take): ReDim chips(1 To take): ReDim blockNos(1 To take)
        For r = 1 To take
            x = AuditInteger(coords(r, 1), -2147483648#, 2147483647#): y = AuditInteger(coords(r, 2), -2147483648#, 2147483647#)
            b = AuditInteger(coords(r, 3), 0, 255)
            chips(r) = x & "|" & y: blockNos(r) = b
        Next r
        For Each g In groups
            flags = ws.Cells(start, CLng(g(0))).Resize(take, CLng(g(1))).Value2
            For j = 1 To CLng(g(1))
                info = columns(CStr(CLng(g(0)) + j - 1))
                activeChip = "": activeTotal = 0
                For r = 1 To take
                    If take = 1 And CLng(g(1)) = 1 Then flag = flags Else flag = flags(r, j)
                    If IsError(flag) Then Err.Raise vbObjectError + 484, , "Address flag cell error."
                    If Len(Trim$(CStr(flag))) > 0 Then
                        If Not IsNumeric(flag) Then Err.Raise vbObjectError + 484, , "Address flag must be0/1 or blank."
                        If CDbl(flag) <> 0 And CDbl(flag) <> 1 Then Err.Raise vbObjectError + 484, , "Address flag must be0/1 or blank."
                        If CDbl(flag) = 1 Then
                            chip = info(0) & "|" & info(1) & "|" & chips(r)
                            If chip <> activeChip Then
                                If Len(activeChip) > 0 Then blocks(activeChip) = bits: totals(activeChip) = activeTotal
                                activeChip = chip: activeTotal = 0
                                If blocks.Exists(chip) Then bits = blocks(chip) Else bits = emptyBits
                                If totals.Exists(chip) Then activeTotal = CLng(totals(chip))
                            End If
                            b = blockNos(r)
                            bits(b \ 8) = CByte(CLng(bits(b \ 8)) Or CLng(2 ^ (b Mod 8)))
                            activeTotal = activeTotal + 1
                        End If
                    End If
                Next r
                If Len(activeChip) > 0 Then blocks(activeChip) = bits: totals(activeChip) = activeTotal
            Next j
        Next g
        Application.StatusBar = "Index Function Address: " & start + take - 1 & "/" & lastRow
    Next start
End Sub

Public Function AuditFunctionBlocks(ByVal wb As Workbook, Optional ByVal outputWorkbook As Workbook = Nothing) As String
    Dim wanted As Object, columns As Object, blocks As Object, totals As Object, seen As Object
    Dim source As Worksheet, stage As Worksheet, old As Worksheet, status As Worksheet, data As Variant, packet() As Variant, rowValues As Variant, info As Variant
    Dim last As Long, start As Long, take As Long, r As Long, b As Long, phase As Long, used As Long, dest As Long, missing As Long, tested As Long
    Dim wt As String, tno As String, chip As String, key As String, prefix As String, x As Long, y As Long, fbc As Double, total As Long, errorNumber As Long
    Dim mappingNote As String, sourceNote As String, description As String, oldAlerts As Boolean, success As Boolean, addressBits As Variant, hasAddress As Boolean, scopeText As String
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    If outputWorkbook Is Nothing Then Set outputWorkbook = wb
    scopeText = Trim$(CStr(outputWorkbook.Worksheets("Config").Range("Q2").Value2))
    If Len(scopeText) = 0 Then scopeText = "All Function"
    Set wanted = AuditFunctions(wb.Worksheets("Config"), CStr(outputWorkbook.Worksheets("Config").Range("Q2").Value2))
    Set columns = ResolveAuditColumns(wb, wanted, mappingNote)
    sourceNote = AuditSourceNote(wb)
    On Error Resume Next
    Set status = wb.Worksheets("Chip Locations")
    On Error GoTo Failed
    If Not status Is Nothing Then
        For phase = 1 To 3
            If CStr(status.Cells(phase + 1, 5).Value2) = "Data Error" Then Err.Raise vbObjectError + 485, , "FBC WT" & phase & " has a parse/output error; fix it before comparison."
        Next phase
    End If
    Set blocks = CreateObject("Scripting.Dictionary"): Set totals = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary")
    BuildDetailedBLKIndex wb.Worksheets("Fail Address"), columns, blocks, totals
    Set source = wb.Worksheets("FBC by BLK")
    last = AuditLastRow(source, 2)
    If last < 2 Then Err.Raise vbObjectError + 485, , "Import FBC before running comparison."
    Set stage = outputWorkbook.Worksheets.Add(After:=outputWorkbook.Worksheets(outputWorkbook.Worksheets.Count))
    stage.Visible = xlSheetVeryHidden
    stage.Range("A4:L4").Value2 = Array("WT", "TNO", "Func.", "Test Voltage", "Pattern", "X", "Y", "BLK", "FBC", "Recorded Chip Addresses", ">=1000 recorded", "Finding")
    dest = 5: ReDim packet(1 To 512, 1 To 12)
    For start = 2 To last Step 512
        take = last - start + 1: If take > 512 Then take = 512
        data = source.Cells(start, 1).Resize(take, 260).Value2
        For r = 1 To take
            If IsError(data(r, 2)) Then Err.Raise vbObjectError + 485, , "Invalid FBC TNO."
            tno = Trim$(CStr(data(r, 2)))
            If wanted.Exists(tno) Then
                wt = UCase$(Trim$(CStr(data(r, 1))))
                If wt <> "WT1" And wt <> "WT2" And wt <> "WT3" Then Err.Raise vbObjectError + 485, , "Invalid FBC WT."
                x = AuditInteger(data(r, 3), -2147483648#, 2147483647#): y = AuditInteger(data(r, 4), -2147483648#, 2147483647#)
                prefix = wt & "|" & tno & "|" & x & "|" & y
                If seen.Exists(prefix) Then Err.Raise vbObjectError + 485, , "Duplicate FBC WT/TNO/Chip record."
                seen.Add prefix, True: info = wanted(tno): tested = tested + 1
                total = 0: If totals.Exists(prefix) Then total = CLng(totals(prefix))
                addressBits = Empty
                If blocks.Exists(prefix) Then addressBits = blocks(prefix)
                For b = 0 To 255
                    fbc = AuditFBCValue(data(r, b + 5))
                    hasAddress = False
                    If IsArray(addressBits) Then hasAddress = ((CLng(addressBits(b \ 8)) And CLng(2 ^ (b Mod 8))) <> 0)
                    If fbc > 0 And Not hasAddress Then
                        missing = missing + 1: used = used + 1
                        If dest + used - 1 > stage.Rows.Count Then Err.Raise vbObjectError + 485, , "Missing BLK output exceeds Excel row limit."
                        packet(used, 1) = wt: packet(used, 2) = CLng(tno): packet(used, 3) = info(0): packet(used, 4) = info(1): packet(used, 5) = info(2)
                        packet(used, 6) = x: packet(used, 7) = y: packet(used, 8) = b: packet(used, 9) = fbc: packet(used, 10) = total
                        packet(used, 11) = IIf(total >= 1000, "Yes", "Not confirmed")
                        packet(used, 12) = "FBC>0; no matching WT/TNO/Chip/BLK Address"
                        If used = 512 Then
                            stage.Cells(dest, 1).Resize(512, 12).Value2 = packet
                            dest = dest + 512: used = 0: ReDim packet(1 To 512, 1 To 12)
                        End If
                    End If
                Next b
            End If
        Next r
        Application.StatusBar = "Compare Function FBC: " & start + take - 1 & "/" & last
    Next start
    If tested = 0 Then Err.Raise vbObjectError + 485, , "FBC contains no Function TNOs in the selected scope."
    If used > 0 Then
        ReDim rowValues(1 To used, 1 To 12)
        For r = 1 To used: For b = 1 To 12: rowValues(r, b) = packet(r, b): Next b: Next r
        stage.Cells(dest, 1).Resize(used, 12).Value2 = rowValues
    End If
    With stage.Cells(1, 1).Resize(missing + 4, 12)
        .Font.Name = "Calibri": .Font.Size = 11: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 26
    End With
    stage.Columns("A:B").ColumnWidth = 10: stage.Columns("C:E").ColumnWidth = 20: stage.Columns("F:H").ColumnWidth = 10
    stage.Columns("I:K").ColumnWidth = 22: stage.Columns("L").ColumnWidth = 56: stage.Columns("D").NumberFormat = "0.000": stage.Columns("I").NumberFormat = "0"
    stage.Range("A1:L1").Merge: stage.Range("A1").Value2 = "Missing Function BLK Address | Scope=" & scopeText & " | TNOs=" & wanted.Count & " | rows=" & missing & " | FBC records=" & tested
    stage.Range("A2:L2").Merge: stage.Range("A2").Value2 = mappingNote & " " & sourceNote
    stage.Range("A2:L2").WrapText = True: stage.Range("A2:L2").RowHeight = 42
    stage.Range("A3:L3").Merge: stage.Range("A3").Value2 = "Source: " & wb.Name & " | Exact WT/TNO/X/Y/BLK match. >=1000 supports a possible cap; missing data alone does not prove the cause."
    With stage.Range("A1:L1"): .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: End With
    With stage.Range("A4:L4"): .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: End With
    stage.Range("A4:L4").Borders.LineStyle = xlContinuous
    stage.Cells(1, 1).Resize(missing + 4, 12).BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(80, 80, 80)
    If missing > 0 Then stage.Range("A5:L" & missing + 4).Interior.Color = RGB(255, 237, 213)
    stage.Range("A4:L" & missing + 4).AutoFilter
    On Error Resume Next
    Set old = outputWorkbook.Worksheets("Missing BLK Address")
    On Error GoTo Failed
    Application.DisplayAlerts = False
    If Not old Is Nothing Then old.Delete
    stage.Name = "Missing BLK Address": stage.Visible = xlSheetVisible: success = True
    stage.Activate: Application.ActiveWindow.DisplayGridlines = False
    Application.DisplayAlerts = oldAlerts
    AuditFunctionBlocks = "Audit complete. Missing Function WT/TNO/Chip/BLK rows: " & missing
    Exit Function
Failed:
    description = Err.Description: errorNumber = Err.Number
    Resume AuditFailureCleanup
AuditFailureCleanup:
    On Error Resume Next
    Application.DisplayAlerts = False
    If Not stage Is Nothing And Not success Then
        stage.Visible = xlSheetVisible
        stage.Delete
    End If
    Application.DisplayAlerts = oldAlerts
    AuditFunctionBlocks = "Audit failed (" & errorNumber & "): " & description
End Function

Private Sub RecordImportIndex(ByVal wb As Workbook, ByVal mode As String, ByRef meta() As Object, ByRef paths() As String, ByVal tnos As Object)
    Dim ws As Worksheet, mapping() As Variant, manifest(1 To 3, 1 To 8) As Variant, headers As Variant, phase As Long, tno As Variant, i As Long, col As Long, start As Long
    Set ws = wb.Worksheets("Import Index")
    If mode = "Address" Then
        ws.Range("A:D").ClearContents
        ws.Range("A1:D1").Value2 = Array("WT", "TNO", "Column", "Header")
        ReDim mapping(1 To tnos.Count * 3, 1 To 4)
        headers = wb.Worksheets("Fail Address").Cells(1, 1).Resize(1, 6 + tnos.Count * 3).Value2
        For phase = 1 To 3
            For Each tno In tnos
                i = i + 1: col = 7 + (phase - 1) * tnos.Count + CLng(tnos(tno))
                mapping(i, 1) = "WT" & phase: mapping(i, 2) = CStr(tno): mapping(i, 3) = col: mapping(i, 4) = headers(1, col)
            Next tno
        Next phase
        ws.Range("A2").Resize(i, 4).Value2 = mapping
    End If
    For phase = 1 To 3
        manifest(phase, 1) = mode: manifest(phase, 2) = "WT" & phase: manifest(phase, 3) = CStr(meta(phase)("LOT"))
        manifest(phase, 4) = NormalWafer(CStr(meta(phase)("WNO"))): manifest(phase, 5) = CStr(meta(phase)("STEP"))
        manifest(phase, 6) = BuildResultBaseName(meta(phase)): manifest(phase, 7) = paths(phase): manifest(phase, 8) = Now
    Next phase
    ws.Range("H1:O1").Value2 = Array("Mode", "WT", "Full LOT", "WNO", "STEP", "Identity", "Source path", "Imported at")
    start = IIf(mode = "Address", 2, 5)
    ws.Cells(start, 8).Resize(3, 8).Value2 = manifest
    ws.Visible = xlSheetVeryHidden
    MarkAuditStale wb, mode
End Sub

Private Sub MarkAuditStale(ByVal wb As Workbook, ByVal mode As String)
    On Error Resume Next
    With wb.Worksheets("Missing BLK Address")
        .Range("A2").Value2 = mode & " source import updated. Run Config J22 to refresh this comparison before using its rows."
        .Range("A2:L2").Interior.Color = RGB(255, 217, 102)
    End With
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
