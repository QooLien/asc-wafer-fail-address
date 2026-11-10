Attribute VB_Name = "ImportFailAddresses"
Option Explicit

Private Const READ_BYTES As Long = 262144
Private Const MAX_PENDING_RECORD As Long = 16384
Private Const MAX_RESULT_ROWS As Long = 1048575
Private mTokenRegex As Object
Private mHeaderRegex As Object
Private mBitMasks As Variant
Private busy As Boolean

Private Sub RestoreProgressStatus(ByVal previous As Variant)
    'Use an empty string so Excel never renders a coerced Boolean as status text.
    If VarType(previous) = vbString Then
        If Len(Trim$(previous)) > 0 And StrComp(Trim$(previous), "False", vbTextCompare) <> 0 Then
            Application.StatusBar = CStr(previous)
            Exit Sub
        End If
    End If
    Application.StatusBar = vbNullString
End Sub

' No per-address logs. Keep one address index and one bit per selected TNO/address.
Public Sub ImportASCFailAddresses()
    RunThreeASC False
End Sub

Public Sub ImportASCAddressOnly()
    RunThreeASC False, True
End Sub

Public Sub ImportASCFBC()
    RunThreeASC True
End Sub

Public Sub BuildFailMaps()
    Dim text As String, started As Date, oldScreen As Boolean, oldEvents As Boolean, oldStatus As Variant
    If busy Then MsgBox "An import or Map build is already running.", vbExclamation: Exit Sub
    busy = True: started = Now
    oldScreen = Application.ScreenUpdating: oldEvents = Application.EnableEvents
    oldStatus = Application.StatusBar
    On Error GoTo FailedMap
    Application.ScreenUpdating = False: Application.EnableEvents = False
    BLKVminEngine.BLK_MigrateReportNames ThisWorkbook
    Application.StatusBar = "Wafer Maps: calculating Fail BLK counts..."
    text = BLKVminEngine.BLK_BuildFailMaps(ThisWorkbook)
    If Len(text) = 0 Then
        text = AuditFunctionBlocks(ThisWorkbook)
        If Left$(text, 15) = "Audit complete." Then text = ""
    End If
    BLKVminEngine.BLK_ApplyPresentation ThisWorkbook
    WriteOperationLog ThisWorkbook, "Map", IIf(Len(text) = 0, "Fail Maps complete.", text), (Len(text) > 0)
    If Len(text) = 0 Then
        Application.StatusBar = "Wafer Maps: updating report layout..."
        BLKVminEngine.BLK_PublishReports ThisWorkbook
    End If
    If Len(text) > 0 Then
        MsgBox text, vbExclamation
    Else
        Application.StatusBar = "Wafer Maps: saving workbook..."
        ThisWorkbook.Save
        Application.StatusBar = "Wafer Maps complete."
        If VarType(Application.Caller) = vbString Then oldEvents = True
        MsgBox "Fail Maps complete. Seconds: " & DateDiff("s", started, Now), vbInformation
    End If
DoneMap:
    RestoreProgressStatus oldStatus
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents: busy = False
    Exit Sub
FailedMap:
    text = "Map failed (" & Err.Number & "): " & Err.Description
    WriteOperationLog ThisWorkbook, "Map", text, True
    MsgBox text, vbExclamation
    Resume DoneMap
End Sub

Private Sub RunThreeASC(ByVal fbcOnly As Boolean, Optional ByVal addressOnly As Boolean = False)
    Dim paths(1 To 3) As String, meta(1 To 3) As Object, steps(1 To 3) As Long
    Dim fileCount As Long, fbcTotal As Long, auditMessage As String
    Dim phase As Long, other As Long, swapPath As String, swapStep As Long, swapMeta As Object
    Dim localTNOs As Object, allAccepted As Long, allSkipped As Long, sourceStats(1 To 3, 1 To 6) As Variant
    Dim baseNames(1 To 3) As String, totalTNOs As Long, color As Long
    Dim blkWarnings As String, blkMessage As String, addressLabels As Object, chipCats As Object
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
    BLKVminEngine.BLK_MigrateReportNames ThisWorkbook
    stage = "Validate Config"
    Application.StatusBar = "Import: checking Config..."
    Set cfg = ThisWorkbook.Worksheets("Config")
    If Not fbcOnly Then Set dst = ThisWorkbook.Worksheets("Fail Address")
    stage = "Validate Monitor / Function settings"
    BLKVminEngine.BLK_Setup cfg
    Set tnos = BLKVminEngine.BLK_AddressTNOs()
    If Not fbcOnly Then Set addressLabels = BuildAddressLabels(cfg)
    stage = "Select ASC files"
    Set fd = Application.FileDialog(3)
    With fd
        .Title = "Select 1 to 3 ASC files": .AllowMultiSelect = True
        .Filters.Clear: .Filters.Add "ASC raw data", "*.asc"
        If .Show <> -1 Then GoTo Done
        fileCount = .SelectedItems.Count
        If fileCount < 1 Or fileCount > 3 Then Err.Raise vbObjectError + 20, , "Select 1 to 3 ASC files."
        For phase = 1 To fileCount: paths(phase) = .SelectedItems(phase): Next phase
    End With
    startedAt = Now
    stage = "Validate LOT, WNO and STEP"
    For phase = 1 To fileCount
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
    For phase = 1 To fileCount - 1
        For other = phase + 1 To fileCount
            If steps(phase) = steps(other) Then Err.Raise vbObjectError + 24, , "Duplicate numeric STEP; cannot assign WT1/WT2/WT3."
            If steps(phase) > steps(other) Then
                swapStep = steps(phase): steps(phase) = steps(other): steps(other) = swapStep
                swapPath = paths(phase): paths(phase) = paths(other): paths(other) = swapPath
                Set swapMeta = meta(phase): Set meta(phase) = meta(other): Set meta(other) = swapMeta
            End If
        Next other
    Next phase
    outputBase = SafeFilePart(Left$(CleanSpace(CStr(meta(1)("LOT"))), 6)) & "#" & NormalWafer(CStr(meta(1)("WNO"))) & "_WT1"
    If fileCount > 1 Then outputBase = outputBase & "-WT" & fileCount
    If addressOnly Then
        outputBase = outputBase & "_Address"
    ElseIf fbcOnly Then
        outputBase = outputBase & "_FBC"
    Else
        outputBase = outputBase & "_All"
    End If
    outputPath = UniqueResultPath(fso, fso.GetParentFolderName(paths(1)), outputBase)
    totalTNOs = fileCount * tnos.Count
    If Not fbcOnly Then
        flagBytes = (totalTNOs + 7) \ 8
        capacity = 4096
        ReDim bits(0 To flagBytes - 1, 0 To capacity - 1)
    End If
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    ThisWorkbook.Worksheets("Import Index").Cells(IIf(fbcOnly, 5, 2), 8).Value2 = IIf(fbcOnly, "FBC Pending", "Address Pending")
    If Not fbcOnly And Not addressOnly Then ThisWorkbook.Worksheets("Import Index").Range("H5").Value2 = "FBC Pending"
    For phase = 1 To fileCount
        path = paths(phase): buffer = "": headerTNO = ""
        selectedSections = 0: accepted = 0: skipped = 0: bytePosition = 0
        Set localTNOs = PhaseTNOs(tnos, phase)
        stage = "Open WT" & phase
        fileNo = FreeFile: Open path For Binary Access Read As #fileNo
        fileOpen = True: fileSize = LOF(fileNo)
        DetectEncoding fileNo, encoding
        BLKVminEngine.BLK_BeginFile path
    Do While Seek(fileNo) <= fileSize
        stage = "Read ASC chunk"
        Application.StatusBar = "Import WT" & phase & " of " & fileCount & ": reading ASC..."
        raw = ReadASCChunk(fileNo, encoding)
        bytePosition = Seek(fileNo) - 1
        If addressOnly Then
            If phase = 1 Then BLKVminEngine.BLK_FeedCategories raw, False
        Else
            BLKVminEngine.BLK_ObserveLocations raw, False
            BLKVminEngine.BLK_Feed raw, False
        End If
        If Not fbcOnly Then buffer = buffer & raw
        raw = ""
        If fbcOnly Then stage = "FBC stream WT" & phase Else stage = "Parse and merge addresses"
        If Not fbcOnly Then ParseASCBuffer buffer, False, headerTNO, selectedSections, accepted, skipped, localTNOs, addresses, bits, flagBytes, capacity
        chunkCount = chunkCount + 1
        If chunkCount Mod 8 = 0 Then
            Application.StatusBar = "Import WT" & phase & " of " & fileCount & ": " & _
                Format$(bytePosition / fileSize, "0%") & " read (" & _
                Format$(bytePosition / 1024, "#,##0") & " / " & Format$(fileSize / 1024, "#,##0") & " KB)"
            DoEvents
        End If
    Loop
        Close #fileNo: fileOpen = False
        stage = "Parse final buffer WT" & phase
        If Not fbcOnly Then ParseASCBuffer buffer, True, headerTNO, selectedSections, accepted, skipped, localTNOs, addresses, bits, flagBytes, capacity
        If Not addressOnly Then
            BLKVminEngine.BLK_ObserveLocations "", True
            BLKVminEngine.BLK_Feed "", True
        End If
        If Not fbcOnly And selectedSections = 0 Then
            If phase = 1 Then Err.Raise vbObjectError + 5, , "WT1: no selected tno= section. File: " & path
            If BLKVminEngine.BLK_AllObservedCatsC20() Then
                WriteOperationLog ThisWorkbook, "WT" & phase, "All recorded TNO 1000 categories are C20. No Fail Address was imported for this WT."
            Else
                WriteOperationLog ThisWorkbook, "WT" & phase, "No selected TNO sections. No Fail Address was imported for this WT."
            End If
        End If
        If Not fbcOnly And accepted = 0 And skipped > 0 Then Err.Raise vbObjectError + 6, , "WT" & phase & ": no valid selected addresses. Previous results retained."
        sourceStats(phase, 1) = "WT" & phase
        sourceStats(phase, 2) = path
        sourceStats(phase, 3) = CStr(meta(phase)("STEP"))
        sourceStats(phase, 4) = BuildResultBaseName(meta(phase))
        sourceStats(phase, 5) = accepted: sourceStats(phase, 6) = skipped
        If addressOnly Then
            If phase = 1 Then
                BLKVminEngine.BLK_FeedCategories "", True
                BLKVminEngine.BLK_SaveCategories ThisWorkbook, 1
            End If
        Else
        stage = "FBC by BLK WT" & phase
        fbcTotal = fbcTotal + BLKVminEngine.BLK_RecordCount()
        If fbcOnly Then accepted = BLKVminEngine.BLK_RecordCount()
        blkMessage = BLKVminEngine.BLK_WritePhase(ThisWorkbook, phase)
        If Len(blkMessage) > 0 Then blkWarnings = blkWarnings & vbCrLf & "WT" & phase & ": " & blkMessage
        End If
        allAccepted = allAccepted + accepted: allSkipped = allSkipped + skipped
    Next phase
    If Not fbcOnly Then
    uniqueCount = addresses.Count
    columnCount = 7 + totalTNOs
    Set chipCats = BLKVminEngine.BLK_Categories(ThisWorkbook, "WT1")
    ReDim headers(1 To 1, 1 To columnCount)
    fields = Array("X", "Y", "BLK", "WL", "BL", "IO")
    For j = 0 To 5: headers(1, j + 1) = fields(j): Next j
    headers(1, 7) = "Cat."
    For phase = 1 To fileCount
        Set localTNOs = PhaseTNOs(tnos, phase)
        For Each v In localTNOs: headers(1, 8 + CLng(localTNOs(v))) = "WT" & phase & vbLf & CStr(addressLabels(CStr(v))): Next v
    Next phase
    stage = "Write output batches"
    Application.StatusBar = "Import: writing results..."
    If dst.AutoFilterMode Then dst.AutoFilterMode = False
    dst.UsedRange.FormatConditions.Delete
    dst.UsedRange.ClearContents
    If dst.UsedRange.Column + dst.UsedRange.Columns.Count - 1 > columnCount Then dst.Range(dst.Columns(columnCount + 1), dst.Columns(dst.UsedRange.Column + dst.UsedRange.Columns.Count - 1)).Delete
    If dst.UsedRange.Row + dst.UsedRange.Rows.Count - 1 > uniqueCount + 1 Then dst.Range(dst.Rows(uniqueCount + 2), dst.Rows(dst.UsedRange.Row + dst.UsedRange.Rows.Count - 1)).Delete
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
            s = fields(0) & "|" & fields(1)
            If chipCats.Exists(s) Then data(batchRow, 7) = chipCats(s)
            For b = 0 To flagBytes - 1
                bitValue = bits(b, rowIndex)
                If bitValue <> 0 Then
                    For k = 0 To 7
                        If (bitValue And mBitMasks(k)) <> 0 Then data(batchRow, 8 + b * 8 + k) = 1
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
    For phase = 1 To fileCount
        Select Case phase
            Case 1: color = RGB(219, 234, 254)
            Case 2: color = RGB(255, 237, 213)
            Case 3: color = RGB(220, 252, 231)
        End Select
        With dst.Cells(1, 8 + (phase - 1) * tnos.Count).Resize(1, tnos.Count)
            .Interior.Color = color: .Font.Color = RGB(23, 50, 77)
            .Orientation = 90: .WrapText = True: .Columns.ColumnWidth = 6
        End With
        If uniqueCount > 0 Then
            With dst.Cells(2, 8 + (phase - 1) * tnos.Count).Resize(uniqueCount, tnos.Count)
                .Interior.Pattern = xlNone
                With .FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=1")
                    .Interior.Color = color
                End With
            End With
        End If
    Next phase

    dst.Cells(1, 1).Resize(uniqueCount + 1, columnCount).AutoFilter
    End If
    If addressOnly Then
        RecordImportIndex ThisWorkbook, "Address", meta, paths, tnos, fileCount
        WriteOperationLog ThisWorkbook, "Address", "Address-only import complete. Unique addresses: " & uniqueCount
        BLKVminEngine.BLK_SizeOutput dst
        BLKVminEngine.BLK_SimplifyButtons ThisWorkbook
        GoTo SaveResult
    End If
    If Not fbcOnly Then RecordImportIndex ThisWorkbook, "Address", meta, paths, tnos, fileCount
    stage = "Build FBC comparison and Maps"
    Application.StatusBar = "Analysis: building FBC comparison and Wafer Maps..."
    blkMessage = BLKVminEngine.BLK_Finalize(ThisWorkbook)
    If Len(blkMessage) > 0 Then
        blkWarnings = blkWarnings & vbCrLf & blkMessage & "; Maps were not updated."
    Else
        blkMessage = BLKVminEngine.BLK_BuildFailMaps(ThisWorkbook)
        If Len(blkMessage) > 0 Then blkWarnings = blkWarnings & vbCrLf & blkMessage
    End If
    WriteOperationLog ThisWorkbook, "FBC", IIf(Len(blkWarnings) = 0, "FBC import complete. Records: " & fbcTotal, blkWarnings), (Len(blkWarnings) > 0)
    WriteOperationLog ThisWorkbook, "Map", IIf(Len(blkWarnings) = 0, "Fail Maps complete.", blkWarnings), (Len(blkWarnings) > 0)
    If Not fbcOnly Then
        WriteOperationLog ThisWorkbook, "Address", "Import complete. Unique addresses: " & uniqueCount
    End If
    RecordImportIndex ThisWorkbook, "FBC", meta, paths, tnos
    If Not fbcOnly And Len(blkWarnings) = 0 Then
        stage = "Audit Map BLK Address coverage"
        Application.StatusBar = "Analysis: checking BLK Address coverage..."
        auditMessage = AuditFunctionBlocks(ThisWorkbook)
        If Left$(auditMessage, 15) <> "Audit complete." Then blkWarnings = blkWarnings & vbCrLf & auditMessage
    End If
    BLKVminEngine.BLK_ApplyPresentation ThisWorkbook
SaveResult:
    Application.StatusBar = "Analysis: preparing final reports..."
    If Not addressOnly And Len(blkWarnings) = 0 Then BLKVminEngine.BLK_PublishReports ThisWorkbook
    BLKVminEngine.BLK_BuildImportInfo ThisWorkbook
    If Not addressOnly And Not fbcOnly And Len(blkWarnings) = 0 Then
        ThisWorkbook.Worksheets("Wafer Maps").Activate
        With Application.ActiveWindow
            .Zoom = 80
            .ScrollRow = 6
            .ScrollColumn = 1
        End With
        ThisWorkbook.Worksheets("Wafer Maps").Range("A6").Select
    End If
    stage = "Save result workbook"
    Application.StatusBar = "Import: saving result workbook..."
    Application.Calculation = oldCalc
    ThisWorkbook.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    If Len(blkWarnings) = 0 Then
        If Not addressOnly And Not fbcOnly Then
            If VarType(Application.Caller) = vbString Then oldEvents = True
        End If
        Application.StatusBar = "Import complete."
    Else
        Application.StatusBar = "Import saved. See Import Info for details."
    End If
    n = DateDiff("s", startedAt, Now)
    elapsedText = Format$(Fix(n / 3600), "00") & ":" & Format$(Fix(n / 60) Mod 60, "00") & ":" & Format$(n Mod 60, "00")
    If addressOnly Then
        MsgBox "Address import complete. Files: " & fileCount & vbCrLf & "Unique addresses: " & uniqueCount & vbCrLf & "Saved as: " & outputPath & vbCrLf & "Run time (hh:mm:ss): " & elapsedText, vbInformation
    ElseIf fbcOnly Then
        MsgBox "FBC import complete. Records: " & fbcTotal & vbCrLf & "Saved as: " & outputPath & vbCrLf & "Run time (hh:mm:ss): " & elapsedText & vbCrLf & "Fail Address was not changed." & blkWarnings, vbInformation
    Else
    MsgBox "Import and analysis complete. Files: " & fileCount & vbCrLf & "Unique addresses: " & uniqueCount & vbCrLf & "FBC records: " & fbcTotal & vbCrLf & "Saved as: " & outputPath & vbCrLf & "Run time (hh:mm:ss): " & elapsedText & vbCrLf & "Total seconds: " & Format$(n, "#,##0") & vbCrLf & "See Import Info for the latest operation message." & blkWarnings, vbInformation
    End If
Done:
    On Error Resume Next
    If fileOpen Then Close #fileNo
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen
    RestoreProgressStatus oldStatus
    BLKVminEngine.BLK_Release
    busy = False
    Exit Sub
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Application.StatusBar = "Import stopped: " & stage
    WriteOperationLog ThisWorkbook, IIf(fbcOnly, "FBC", "Address"), "Import failed (" & errorNumber & "): " & errorText & vbCrLf & "Stage: " & stage & "; bytes read: " & bytePosition, True
    MsgBox "Import failed (" & errorNumber & "): " & errorText & vbCrLf & "Stage: " & stage & "; bytes read: " & bytePosition, vbExclamation
    Resume Done
End Sub


Public Function MapChipAddressEvidence(ByVal wb As Workbook, Optional ByVal functionOnly As Boolean = False) As Object
    Dim wantedFunctions As Object, allowed As Boolean
    Dim result As Object, ws As Worksheet, idx As Worksheet, data As Variant, headers As Variant
    Dim xy As Variant, flags As Variant, phase As Long, i As Long, j As Long, r As Long
    Dim firstCol(1 To 3) As Long, lastCol(1 To 3) As Long, last As Long, lastRow As Long
    Dim start As Long, take As Long, batch As Long, col As Long, wt As String, key As String
    Set result = CreateObject("Scripting.Dictionary")
    Set MapChipAddressEvidence = result
    If functionOnly Then Set wantedFunctions = AuditFunctions(wb.Worksheets("Config"), "ALL FUNCTION")
    Set ws = wb.Worksheets("Fail Address"): Set idx = wb.Worksheets("Import Index")
    last = idx.Cells(idx.Rows.Count, 1).End(xlUp).Row
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last < 2 Or lastRow < 2 Then Exit Function
    headers = ws.Range(ws.Cells(1, 1), ws.Cells(1, ws.Columns.Count).End(xlToLeft)).Value2
    data = idx.Range("A2:D" & last).Value2
    For i = 1 To UBound(data, 1)
        wt = CStr(data(i, 1))
        allowed = True
        If functionOnly Then allowed = wantedFunctions.Exists(CStr(data(i, 2)))
        If allowed And (wt = "WT1" Or wt = "WT2" Or wt = "WT3") Then
            phase = CLng(Right$(wt, 1)): col = CLng(data(i, 3))
            If col < 8 Or col > UBound(headers, 2) Then Err.Raise vbObjectError + 493, , "Address evidence index out of range."
            If CStr(headers(1, col)) <> CStr(data(i, 4)) Then Err.Raise vbObjectError + 493, , "Address evidence header changed; reimport Address."
            If firstCol(phase) = 0 Or col < firstCol(phase) Then firstCol(phase) = col
            If col > lastCol(phase) Then lastCol(phase) = col
        End If
    Next i
    For phase = 1 To 3
        If firstCol(phase) > 0 Then
            batch = 131072 \ (lastCol(phase) - firstCol(phase) + 1): If batch > 4096 Then batch = 4096
            For start = 2 To lastRow Step batch
                take = lastRow - start + 1: If take > batch Then take = batch
                xy = ws.Cells(start, 1).Resize(take, 2).Value2
                flags = ws.Cells(start, firstCol(phase)).Resize(take, lastCol(phase) - firstCol(phase) + 1).Value2
                For r = 1 To take
                    key = "WT" & phase & "|" & CStr(xy(r, 1)) & "|" & CStr(xy(r, 2))
                    If Not result.Exists(key) Then
                        If Not IsArray(flags) Then
                            If flags = 1 Then result.Add key, True
                        Else
                            For j = 1 To UBound(flags, 2)
                                If Not IsError(flags(r, j)) Then
                                    If flags(r, j) = 1 Then result.Add key, True: Exit For
                                End If
                            Next j
                        End If
                    End If
                Next r
            Next start
        End If
    Next phase
End Function

Public Function BuildAddressLabels(ByVal cfg As Worksheet) As Object
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
    Application.StatusBar = "Audit: checking BLK Address coverage..."
    result = AuditFunctionBlocks(source, ThisWorkbook)
    If Left$(result, 12) = "Audit failed" Then
        MsgBox result, vbExclamation
    Else
        Application.StatusBar = "Audit: saving results..."
        ThisWorkbook.Save
        Application.StatusBar = "Audit complete."
        MsgBox result & vbCrLf & "Seconds: " & DateDiff("s", started, Now), vbInformation
    End If
DoneAudit:
    On Error Resume Next
    If opened Then source.Close SaveChanges:=False
    Application.AutomationSecurity = oldSecurity
    Application.ScreenUpdating = oldScreen: Application.EnableEvents = oldEvents: Application.Calculation = oldCalc
    RestoreProgressStatus oldStatus
    busy = False: Exit Sub
FailedAudit:
    MsgBox "Audit failed (" & Err.Number & "): " & Err.Description, vbExclamation
    Resume DoneAudit
End Sub

Private Function AuditFunctions(ByVal cfg As Worksheet, Optional ByVal scopeOverride As String = "") As Object
    Dim result As Object, data As Variant, i As Long, tno As String, stopOnly As Boolean, scope As String, func As String, voltage As Variant, pattern As String
    Set result = CreateObject("Scripting.Dictionary")
    If Len(scopeOverride) > 0 Then scope = UCase$(Trim$(scopeOverride)) Else scope = UCase$(Trim$(CStr(cfg.Range("Q2").Value2)))
    If scope <> "" And scope <> "ALL FUNCTION" And scope <> "STOP SIGN" And scope <> "FAIL MAP" Then Err.Raise vbObjectError + 480, , "Audit scope Q2 must be All Function or Stop Sign."
    stopOnly = (scope = "STOP SIGN" Or scope = "FAIL MAP")
    data = cfg.Range("D2:G901").Value2
    For i = LBound(data, 1) To UBound(data, 1)
        If IsError(data(i, 1)) Or IsError(data(i, 4)) Then Err.Raise vbObjectError + 480, , "Function configuration error at row " & i + 1
        tno = Trim$(CStr(data(i, 1))): pattern = Trim$(CStr(data(i, 4)))
        If Len(tno) > 0 And (Not stopOnly Or Len(pattern) > 0) And (scope <> "FAIL MAP" Or StrComp(pattern, "Hard Fail", vbTextCompare) <> 0) Then
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
        For phase = 1 To BLKVminEngine.BLK_PhaseCount(wb)
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
    Dim blockNos() As Long, bits As Variant, activeChip As String, activeTotal As Long, emptyBits(0 To 255) As Long
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
                            bits(b) = CLng(bits(b)) + 1
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

Private Function BuildPatternAddressIndex(ByVal wb As Workbook, ByVal endpoints As Object, ByVal columns As Object) As Object
    Dim result As Object, lookup As Object, scope As Object, chosen As Collection, ws As Worksheet
    Dim endpoint As Variant, entry As Variant, info As Variant, tno As Variant, key As Variant, flag As Variant
    Dim col As Long, phase As Long, low As Long, high As Long, firstRow As Long, lastConfigRow As Long
    Dim lastRow As Long, start As Long, take As Long, batch As Long, r As Long, matched As Boolean
    Dim xy As Variant, flags As Variant, bits As Variant, emptyBits(0 To 255) As Long, chip As String, block As Long, activeChip As String
    Set result = CreateObject("Scripting.Dictionary"): Set lookup = CreateObject("Scripting.Dictionary")
    For Each key In columns
        info = columns(key): lookup(CStr(info(0)) & "|" & CStr(info(1))) = CLng(key)
    Next key
    Set ws = wb.Worksheets("Fail Address"): lastRow = AuditLastRow(ws, 1)
    For Each endpoint In endpoints
        entry = endpoints(endpoint)
        Set scope = DrillFunctionScope(wb.Worksheets("Config"), CStr(endpoint), CStr(entry(2)), firstRow, lastConfigRow)
        For phase = 1 To BLKVminEngine.BLK_PhaseCount(wb)
            If CStr(wb.Worksheets("Chip Locations").Cells(phase + 1, 5).Value2) = "No Data" Then GoTo NextIntervalWT
            Set chosen = New Collection: low = ws.Columns.Count: high = 0
            For Each tno In scope
                key = "WT" & phase & "|" & CStr(tno)
                If Not lookup.Exists(key) Then Err.Raise vbObjectError + 494, , "Pattern interval Address column unavailable: " & key
                col = CLng(lookup(key)): chosen.Add col
                If col < low Then low = col
                If col > high Then high = col
            Next tno
            If chosen.Count = 0 Then Err.Raise vbObjectError + 494, , "Empty Pattern interval."
            batch = 131072 \ (high - low + 1): If batch > 4096 Then batch = 4096
            For start = 2 To lastRow Step batch
                take = lastRow - start + 1: If take > batch Then take = batch
                xy = ws.Cells(start, 1).Resize(take, 3).Value2
                flags = ws.Cells(start, low).Resize(take, high - low + 1).Value2
                activeChip = ""
                For r = 1 To take
                    matched = False
                    For Each tno In chosen
                        If take = 1 And high = low Then flag = flags Else flag = flags(r, CLng(tno) - low + 1)
                        If IsError(flag) Then Err.Raise vbObjectError + 494, , "Pattern interval Address flag cell error."
                        If Len(Trim$(CStr(flag))) > 0 Then
                            If Not IsNumeric(flag) Then Err.Raise vbObjectError + 494, , "Pattern interval flag must be 0/1 or blank."
                            If CDbl(flag) <> 0 And CDbl(flag) <> 1 Then Err.Raise vbObjectError + 494, , "Pattern interval flag must be 0/1 or blank."
                            If CDbl(flag) = 1 Then matched = True: Exit For
                        End If
                    Next tno
                    If matched Then
                        chip = "WT" & phase & "|" & CStr(endpoint) & "|" & CStr(xy(r, 1)) & "|" & CStr(xy(r, 2))
                        block = AuditInteger(xy(r, 3), 0, 255)
                        If chip <> activeChip Then
                            If Len(activeChip) > 0 Then result(activeChip) = bits
                            activeChip = chip
                            If result.Exists(chip) Then bits = result(chip) Else bits = emptyBits
                        End If
                        bits(block) = CLng(bits(block)) + 1
                    End If
                Next r
                If Len(activeChip) > 0 Then result(activeChip) = bits
            Next start
NextIntervalWT:
        Next phase
    Next endpoint
    Set BuildPatternAddressIndex = result
End Function

Public Function AuditFunctionBlocks(ByVal wb As Workbook, Optional ByVal outputWorkbook As Workbook = Nothing) As String
    Dim wanted As Object, columns As Object, blocks As Object, totals As Object, seen As Object, mapMasks As Object, allFunctions As Object, patternBlocks As Object, patternBits As Variant, auditCol As Variant
    Dim source As Worksheet, stage As Worksheet, old As Worksheet, status As Worksheet, data As Variant, packet() As Variant, rowValues As Variant, info As Variant
    Dim last As Long, start As Long, take As Long, r As Long, b As Long, phase As Long, used As Long, dest As Long, missing As Long, tested As Long
    Dim wt As String, tno As String, chip As String, key As String, prefix As String, x As Long, y As Long, fbc As Double, total As Long, errorNumber As Long
    Dim mappingNote As String, sourceNote As String, description As String, oldAlerts As Boolean, success As Boolean, addressBits As Variant, hasAddress As Boolean, scopeText As String
    Dim recorded As Long, gap As Double, included As String, finding As String
    Dim completeCount As Long, partialCount As Long, absentCount As Long, excessCount As Long
    Dim intervalCount As Long, endpointFinding As String, coverageBasis As String, intervalOnlyCount As Long, monitorNote As String
    oldAlerts = Application.DisplayAlerts
    On Error GoTo Failed
    If outputWorkbook Is Nothing Then Set outputWorkbook = wb
    If outputWorkbook Is wb Then BLKVminEngine.BLK_InvalidateCoverageMap wb
    scopeText = "Fail Map included BLKs"
    Set wanted = AuditFunctions(wb.Worksheets("Config"), "FAIL MAP")
    Set allFunctions = AuditFunctions(wb.Worksheets("Config"), "ALL FUNCTION")
    Set columns = ResolveAuditColumns(wb, allFunctions, mappingNote)
    sourceNote = AuditSourceNote(wb)
    On Error Resume Next
    Set status = wb.Worksheets("Chip Locations")
    On Error GoTo Failed
    If Not status Is Nothing Then
        For phase = 1 To 3
            If CStr(status.Cells(phase + 1, 5).Value2) = "Data Error" Then Err.Raise vbObjectError + 485, , "FBC WT" & phase & " has a parse/output error; fix it before comparison."
        Next phase
    End If
    Set mapMasks = BLKVminEngine.BLK_MapIncludedMasks(wb)
    Set blocks = CreateObject("Scripting.Dictionary"): Set totals = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary")
    Set patternBlocks = BuildPatternAddressIndex(wb, wanted, columns)
    For Each auditCol In columns.Keys
        rowValues = columns(auditCol)
        If Not wanted.Exists(CStr(rowValues(1))) Then columns.Remove auditCol
    Next auditCol
    BuildDetailedBLKIndex wb.Worksheets("Fail Address"), columns, blocks, totals
    Set source = wb.Worksheets("FBC by BLK")
    last = AuditLastRow(source, 2)
    If last < 2 Then Err.Raise vbObjectError + 485, , "Import FBC before running comparison."
    Set stage = outputWorkbook.Worksheets.Add(After:=outputWorkbook.Worksheets(outputWorkbook.Worksheets.Count))
    stage.Visible = xlSheetVeryHidden
    stage.Range("A4:O4").Value2 = Array("WT", "TNO", "Func.", "Test Voltage", "Pattern", "X", "Y", "BLK", "FBC", "Endpoint Address Count", "Endpoint Missing Count", "Finding", "Pattern Address Count", "Coverage Basis", "Endpoint Finding")
    dest = 5: ReDim packet(1 To 512, 1 To 15)
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
                included = ""
                If mapMasks.Exists(prefix) Then included = CStr(mapMasks(prefix))
                addressBits = Empty
                If blocks.Exists(prefix) Then addressBits = blocks(prefix)
                patternBits = Empty
                If patternBlocks.Exists(prefix) Then patternBits = patternBlocks(prefix)
                For b = 0 To 255
                    fbc = AuditFBCValue(data(r, b + 5))
                    recorded = 0
                    If IsArray(addressBits) Then recorded = CLng(addressBits(b))
                    If fbc > 0 And Mid$(included, b + 1, 1) = "1" Then
                        gap = fbc - recorded: If gap < 0 Then gap = 0
                        If recorded = 0 Then
                            endpointFinding = "No Address"
                        ElseIf recorded < fbc Then
                            endpointFinding = "Partial Address": partialCount = partialCount + 1
                        ElseIf recorded = fbc Then
                            endpointFinding = "Complete"
                        Else
                            endpointFinding = "Address Count > FBC": excessCount = excessCount + 1
                        End If
                        intervalCount = 0
                        If IsArray(patternBits) Then intervalCount = CLng(patternBits(b))
                        If intervalCount > 0 Then
                            finding = "Complete": completeCount = completeCount + 1
                            If recorded > 0 Then
                                coverageBasis = "Endpoint Address present"
                            Else
                                coverageBasis = "Earlier TNO only": intervalOnlyCount = intervalOnlyCount + 1
                            End If
                        Else
                            finding = "No Address": absentCount = absentCount + 1
                            coverageBasis = "No Address in Pattern"
                        End If
                        missing = missing + 1: used = used + 1
                        If dest + used - 1 > stage.Rows.Count Then Err.Raise vbObjectError + 485, , "Missing BLK output exceeds Excel row limit."
                        packet(used, 1) = wt: packet(used, 2) = CLng(tno): packet(used, 3) = info(0): packet(used, 4) = info(1): packet(used, 5) = info(2)
                        packet(used, 6) = x: packet(used, 7) = y: packet(used, 8) = b: packet(used, 9) = fbc: packet(used, 10) = recorded
                        packet(used, 11) = gap
                        packet(used, 12) = finding
                        packet(used, 13) = intervalCount: packet(used, 14) = coverageBasis: packet(used, 15) = endpointFinding
                        If used = 512 Then
                            stage.Cells(dest, 1).Resize(512, 15).Value2 = packet
                            dest = dest + 512: used = 0: ReDim packet(1 To 512, 1 To 15)
                        End If
                    End If
                Next b
            End If
        Next r
        Application.StatusBar = "Compare Function FBC: " & start + take - 1 & "/" & last
    Next start
    If tested = 0 Then Err.Raise vbObjectError + 485, , "FBC contains no Function TNOs in the selected scope."
    If used > 0 Then
        ReDim rowValues(1 To used, 1 To 15)
        For r = 1 To used: For b = 1 To 15: rowValues(r, b) = packet(r, b): Next b: Next r
        stage.Cells(dest, 1).Resize(used, 15).Value2 = rowValues
    End If
    With stage.Cells(1, 1).Resize(missing + 4, 15)
        .Font.Name = "Calibri": .Font.Size = 11: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 26
    End With
    stage.Columns("A:B").ColumnWidth = 10: stage.Columns("C:E").ColumnWidth = 20: stage.Columns("F:H").ColumnWidth = 10
    stage.Columns("I:K").ColumnWidth = 22: stage.Columns("L").ColumnWidth = 20: stage.Columns("M").ColumnWidth = 24: stage.Columns("N:O").ColumnWidth = 30: stage.Columns("D").NumberFormat = "0.000": stage.Columns("I").NumberFormat = "0"
    stage.Range("A1:O1").Merge: stage.Range("A1").Value2 = "Map BLK Address Coverage | Complete=" & completeCount & " | Endpoint Partial=" & partialCount & " | No Address=" & absentCount & " | Endpoint Excess=" & excessCount & " | Earlier only=" & intervalOnlyCount & " | rows=" & missing
    stage.Range("A2:O2").Merge: stage.Range("A2").Value2 = mappingNote & " " & sourceNote
    stage.Range("A2:O2").WrapText = True: stage.Range("A2:O2").RowHeight = 42
    stage.Range("A3:O3").Merge: stage.Range("A3").Value2 = "Source: " & wb.Name & " | Map-counted BLKs; Complete = any Address in same WT/Pattern/Chip/BLK. J/K/O retain exact endpoint diagnostics; earlier-only does not prove output cap."
    With stage.Range("A1:O1"): .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: End With
    With stage.Range("A4:O4"): .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: End With
    stage.Range("A4:O4").Borders.LineStyle = xlContinuous
    stage.Cells(1, 1).Resize(missing + 4, 15).BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(80, 80, 80)
    If missing > 0 Then
        With stage.Range("A5:O" & missing + 4)
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$L5=""Complete"""
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(220, 252, 231)
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$L5=""Partial Address"""
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 237, 213)
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$L5=""No Address"""
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(254, 202, 202)
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$L5=""Address Count > FBC"""
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 217, 102)
        End With
    End If
    If missing > 0 Then
        With stage.Range("N5:N" & missing + 4)
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$N5=""Earlier TNO only"""
            With .FormatConditions(.FormatConditions.Count)
                .Interior.Color = RGB(219, 234, 254): .Font.Color = RGB(23, 50, 77): .SetFirstPriority
            End With
        End With
    End If
    stage.Range("A4:O" & missing + 4).AutoFilter
    On Error Resume Next
    Set old = outputWorkbook.Worksheets("Missing BLK Address")
    On Error GoTo Failed
    Application.DisplayAlerts = False
    If Not old Is Nothing Then old.Delete
    stage.Name = "Missing BLK Address": stage.Visible = xlSheetVisible: success = True
    BLKVminEngine.BLK_SizeOutput stage
    stage.Activate: Application.ActiveWindow.DisplayGridlines = False
    Application.DisplayAlerts = oldAlerts
    If outputWorkbook Is wb Then BLKVminEngine.BLK_BuildCoverageMap wb
    monitorNote = BuildMonitorFailAnalysis(wb, outputWorkbook, mapMasks)
    If Left$(monitorNote, 26) <> "Monitor analysis complete." Then Err.Raise vbObjectError + 495, , monitorNote
    AuditFunctionBlocks = "Audit complete. Map BLKs=" & missing & "; Complete=" & completeCount & "; Endpoint Partial=" & partialCount & "; No Address=" & absentCount & "; Endpoint Excess=" & excessCount & "; Earlier only=" & intervalOnlyCount
    AuditFunctionBlocks = AuditFunctionBlocks & "; " & monitorNote
    WriteOperationLog outputWorkbook, "Address Coverage", AuditFunctionBlocks
    BLKVminEngine.BLK_OrderSheets outputWorkbook
    Exit Function
Failed:
    description = Err.Description: errorNumber = Err.Number
    Resume AuditFailureCleanup
AuditFailureCleanup:
    On Error Resume Next
    If outputWorkbook Is wb Then BLKVminEngine.BLK_InvalidateCoverageMap wb
    Application.DisplayAlerts = False
    If Not stage Is Nothing And Not success Then
        stage.Visible = xlSheetVisible
        stage.Delete
    End If
    Application.DisplayAlerts = oldAlerts
    AuditFunctionBlocks = "Audit failed (" & errorNumber & "): " & description
    If Not outputWorkbook Is Nothing Then WriteOperationLog outputWorkbook, "Address Coverage", AuditFunctionBlocks, True
End Function

Private Function MonitorKind(ByVal text As String) As String
    text = UCase$(Replace(Replace(Replace(Trim$(text), " ", ""), "-", ""), "_", ""))
    If text = "WLOWPBL" Then MonitorKind = "W"
    If text = "RLOWPBL" Then MonitorKind = "R"
End Function

Private Function MonitorOutputSheet(ByVal wb As Workbook, ByVal name As String) As Worksheet
    Dim ws As Worksheet, last As Long
    On Error Resume Next
    Set ws = wb.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count)): ws.Name = name
    ws.Visible = xlSheetVisible
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    last = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    ws.UsedRange.UnMerge: ws.UsedRange.Clear
    If last > 4 Then ws.Rows(5).Resize(last - 4).Delete
    Set MonitorOutputSheet = ws
End Function

Private Sub WriteMonitorPacket(ByVal ws As Worksheet, ByRef packet() As Variant, ByRef used As Long, ByRef dest As Long)
    Dim tail() As Variant, i As Long, j As Long
    If used = 0 Then Exit Sub
    If dest + used - 1 > ws.Rows.Count Then Err.Raise vbObjectError + 495, , "Monitor output exceeds Excel row limit."
    If used = UBound(packet, 1) Then
        ws.Cells(dest, 1).Resize(used, 16).Value2 = packet
    Else
        ReDim tail(1 To used, 1 To 16)
        For i = 1 To used: For j = 1 To 16: tail(i, j) = packet(i, j): Next j: Next i
        ws.Cells(dest, 1).Resize(used, 16).Value2 = tail
    End If
    dest = dest + used: used = 0: ReDim packet(1 To 512, 1 To 16)
End Sub

Private Function MonitorFlag(ByRef flags As Variant, ByVal row As Long, ByVal offset As Long) As Boolean
    Dim flag As Variant
    If IsArray(flags) Then flag = flags(row, offset) Else flag = flags
    If IsError(flag) Then Err.Raise vbObjectError + 495, , "Monitor matching Address flag error."
    If Len(Trim$(CStr(flag))) = 0 Then Exit Function
    If Not IsNumeric(flag) Then Err.Raise vbObjectError + 495, , "Monitor matching flag must be 0/1 or blank."
    If CDbl(flag) <> 0 And CDbl(flag) <> 1 Then Err.Raise vbObjectError + 495, , "Monitor matching flag must be 0/1 or blank."
    MonitorFlag = (CDbl(flag) = 1)
End Function

Private Sub FormatMonitorOutput(ByVal ws As Worksheet, ByVal last As Long)
    Dim c As Long
    With ws.Cells(1, 1).Resize(last, 16)
        .Font.Name = "Calibri": .Font.Size = 11: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .RowHeight = 26
    End With
    ws.Columns("A:P").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 18
    ws.Rows(1).RowHeight = 30: ws.Rows(2).RowHeight = 42: ws.Rows(3).RowHeight = 30
    ws.Range("A2:P2").WrapText = True
    With ws.Range("A4:P4")
        .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True
        .Orientation = 90: .WrapText = True: .RowHeight = 160: .Borders.LineStyle = xlContinuous
    End With
    With ws.Range("A1:P1"): .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: End With
    If ws.Name = "Monitor Fail Analysis" Then
        ws.Columns("J:L").ColumnWidth = 24: ws.Columns("O").ColumnWidth = 28
        If last >= 5 Then
            ws.Range("J5:L" & last).WrapText = True
            ws.Cells(5, 1).Resize(last - 4, 16).RowHeight = 60
        End If
    Else
        ws.Columns("O:P").ColumnWidth = 28
    End If
    ws.Range("A4:P" & last).AutoFilter
    ws.Cells(1, 1).Resize(last, 16).BorderAround LineStyle:=xlContinuous, Weight:=xlThin
    ws.Activate: Application.ActiveWindow.DisplayGridlines = False
End Sub

Public Function BuildMonitorFailAnalysis(ByVal wb As Workbook, Optional ByVal outputWorkbook As Workbook = Nothing, Optional ByVal mapMasks As Object = Nothing) As String
    Dim wanted As Object, functions As Object, monitors As Object, columns As Object, lookup As Object
    Dim endpoints As Object, scopes As Object, scope As Object, catalog As Object, chipCounts As Object
    Dim detail As Worksheet, summary As Worksheet, report As Worksheet, source As Worksheet, cfg As Worksheet
    Dim data As Variant, xy As Variant, flags As Variant, info As Variant, entry As Variant, item As Variant
    Dim packet() As Variant, sums() As Variant, endpoint As Variant, tno As Variant, colKey As Variant
    Dim writeCols As Collection, readCols As Collection, chosen As Collection, group As Variant, groups As Collection
    Dim low As Long, high As Long, col As Long, phase As Long, last As Long, first As Long, take As Long, i As Long, j As Long
    Dim r As Long, start As Long, batch As Long, used As Long, dest As Long, summaryUsed As Long, summaryDest As Long
    Dim firstConfig As Long, lastConfig As Long, code As Long, total As Long, wCount As Long, readCount As Long
    Dim key As String, chipKey As String, prefix As String, mask As String, wt As String, kind As String, mappingNote As String, sourceNote As String
    Dim wList As String, rList As String, functionList As String, label As String, errorText As String
    On Error GoTo Failed
    If outputWorkbook Is Nothing Then Set outputWorkbook = wb
    Set detail = MonitorOutputSheet(outputWorkbook, "Monitor Fail Analysis")
    Set summary = MonitorOutputSheet(outputWorkbook, "Monitor BLK Summary")
    Set cfg = wb.Worksheets("Config"): Set source = wb.Worksheets("Fail Address")
    Set report = outputWorkbook.Worksheets("Missing BLK Address")
    If CStr(report.Range("M4").Value2) <> "Pattern Address Count" Then Err.Raise vbObjectError + 495, , "Refresh Pattern Address coverage before Monitor analysis."
    Set functions = AuditFunctions(cfg, "ALL FUNCTION"): Set endpoints = AuditFunctions(cfg, "FAIL MAP")
    Set wanted = CreateObject("Scripting.Dictionary"): Set monitors = CreateObject("Scripting.Dictionary")
    For Each tno In functions: wanted.Add CStr(tno), True: Next tno
    data = cfg.Range("A2:C901").Value2
    For i = 1 To UBound(data, 1)
        If Len(Trim$(CStr(data(i, 1)))) > 0 Then
            kind = MonitorKind(CStr(data(i, 2)))
            If Len(kind) > 0 Then
                tno = CStr(AuditInteger(data(i, 1), 100, 999))
                If wanted.Exists(tno) Then Err.Raise vbObjectError + 495, , "Duplicate Monitor/Function TNO: " & tno
                wanted.Add CStr(tno), True: monitors.Add CStr(tno), kind
                If kind = "W" Then wCount = wCount + 1 Else readCount = readCount + 1
            End If
        End If
    Next i
    If wCount = 0 Or readCount = 0 Then Err.Raise vbObjectError + 495, , "Configure both W-Low pBL and R-Low pBL Monitor items in Config A:C."
    Set columns = ResolveAuditColumns(wb, wanted, mappingNote): sourceNote = AuditSourceNote(wb)
    Set lookup = CreateObject("Scripting.Dictionary")
    For Each colKey In columns
        info = columns(colKey): lookup(CStr(info(0)) & "|" & CStr(info(1))) = CLng(colKey)
    Next colKey
    Set scopes = CreateObject("Scripting.Dictionary")
    For Each endpoint In endpoints
        info = endpoints(endpoint)
        Set scope = DrillFunctionScope(cfg, CStr(endpoint), CStr(info(2)), firstConfig, lastConfig)
        scopes.Add CStr(endpoint), scope
    Next endpoint
    If mapMasks Is Nothing Then Set mapMasks = BLKVminEngine.BLK_MapIncludedMasks(wb)
    Set catalog = CreateObject("Scripting.Dictionary"): Set chipCounts = CreateObject("Scripting.Dictionary")
    last = report.Cells(report.Rows.Count, 1).End(xlUp).Row
    For first = 5 To last Step 512
        take = last - first + 1: If take > 512 Then take = 512
        data = report.Cells(first, 1).Resize(take, 15).Value2
        For i = 1 To take
            'Coverage also includes Hard Fail and other loss-only rows. Monitor is scoped to Map-counted endpoint BLKs.
            tno = CStr(data(i, 2))
            If endpoints.Exists(tno) Then
                prefix = CStr(data(i, 1)) & "|" & tno & "|" & CStr(data(i, 6)) & "|" & CStr(data(i, 7))
                If mapMasks.Exists(prefix) Then
                    mask = CStr(mapMasks(prefix))
                    If Mid$(mask, CLng(data(i, 8)) + 1, 1) = "1" Then
                        key = prefix & "|" & CStr(data(i, 8))
                        If catalog.Exists(key) Then Err.Raise vbObjectError + 495, , "Duplicate Map BLK in Monitor catalog."
                        'Entry: WT,Pattern,StopTNO,X,Y,BLK,FBC,expected interval count,basis,endpoint gap,observed,W,R,WR,None.
                        catalog.Add key, Array(data(i, 1), data(i, 5), data(i, 2), data(i, 6), data(i, 7), data(i, 8), data(i, 9), data(i, 13), data(i, 14), data(i, 11), 0&, 0&, 0&, 0&, 0&)
                        chipKey = CStr(data(i, 1)) & "|" & CStr(data(i, 5)) & "|" & CStr(data(i, 6)) & "|" & CStr(data(i, 7))
                        If Not chipCounts.Exists(chipKey) Then chipCounts.Add chipKey, 0
                        chipCounts(chipKey) = CLng(chipCounts(chipKey)) + 1
                    End If
                End If
            End If
        Next i
    Next first
    detail.Range("A4:P4").Value2 = Array("WT", "Pattern", "Stop TNO", "X", "Y", "BLK", "WL", "BL", "IO", "Function Fail TNOs", "Write Monitor TNOs", "Read Monitor TNOs", "Write Monitor Fail", "Read Monitor Fail", "Classification", "Source Address Row")
    summary.Range("A4:P4").Value2 = Array("WT", "Pattern", "Stop TNO", "X", "Y", "Chip Fail BLK Count", "BLK", "FBC", "Pattern Address Count", "Write Fail", "Read Fail", "Write&Read Fail", "No monitor item Fail", "Endpoint Missing Count", "Coverage Basis", "Analysis Status")
    detail.Columns("J:L").NumberFormat = "@"
    dest = 5: ReDim packet(1 To 512, 1 To 16)
    last = AuditLastRow(source, 1)
    For phase = 1 To BLKVminEngine.BLK_PhaseCount(wb)
        wt = "WT" & phase: Set writeCols = New Collection: Set readCols = New Collection: Set groups = New Collection
        low = source.Columns.Count: high = 0
        For Each tno In monitors
            col = CLng(lookup(wt & "|" & CStr(tno)))
            If monitors(tno) = "W" Then writeCols.Add Array(col, CStr(tno)) Else readCols.Add Array(col, CStr(tno))
            If col < low Then low = col
            If col > high Then high = col
        Next tno
        For Each endpoint In endpoints
            Set chosen = New Collection: Set scope = scopes(endpoint)
            For Each tno In scope
                col = CLng(lookup(wt & "|" & CStr(tno))): chosen.Add Array(col, CStr(tno))
                If col < low Then low = col
                If col > high Then high = col
            Next tno
            info = endpoints(endpoint): groups.Add Array(CStr(endpoint), CStr(info(2)), chosen)
        Next endpoint
        batch = 131072 \ (high - low + 1): If batch > 4096 Then batch = 4096
        For start = 2 To last Step batch
            take = last - start + 1: If take > batch Then take = batch
            xy = source.Cells(start, 1).Resize(take, 6).Value2
            flags = source.Cells(start, low).Resize(take, high - low + 1).Value2
            For r = 1 To take
                wList = "": rList = ""
                For Each item In writeCols
                    If MonitorFlag(flags, r, CLng(item(0)) - low + 1) Then
                        If Len(wList) > 0 Then wList = wList & ","
                        wList = wList & CStr(item(1))
                    End If
                Next item
                For Each item In readCols
                    If MonitorFlag(flags, r, CLng(item(0)) - low + 1) Then
                        If Len(rList) > 0 Then rList = rList & ","
                        rList = rList & CStr(item(1))
                    End If
                Next item
                For Each group In groups
                    key = wt & "|" & CStr(group(0)) & "|" & CStr(xy(r, 1)) & "|" & CStr(xy(r, 2)) & "|" & CStr(xy(r, 3))
                    If catalog.Exists(key) Then
                        functionList = "": Set chosen = group(2)
                        For Each item In chosen
                            If MonitorFlag(flags, r, CLng(item(0)) - low + 1) Then
                                If Len(functionList) > 0 Then functionList = functionList & ","
                                functionList = functionList & CStr(item(1))
                            End If
                        Next item
                        If Len(functionList) > 0 Then
                            code = 14: label = "No monitor item Fail"
                            If Len(wList) > 0 And Len(rList) > 0 Then
                                code = 13: label = "Write&Read Fail"
                            ElseIf Len(wList) > 0 Then
                                code = 11: label = "Write Fail"
                            ElseIf Len(rList) > 0 Then
                                code = 12: label = "Read Fail"
                            End If
                            entry = catalog(key): entry(10) = CLng(entry(10)) + 1: entry(code) = CLng(entry(code)) + 1: catalog(key) = entry
                            used = used + 1: total = total + 1
                            packet(used, 1) = wt: packet(used, 2) = group(1): packet(used, 3) = CLng(group(0))
                            For j = 1 To 6: packet(used, j + 3) = xy(r, j): Next j
                            packet(used, 10) = functionList: packet(used, 11) = wList: packet(used, 12) = rList
                            packet(used, 13) = IIf(Len(wList) > 0, 1, 0): packet(used, 14) = IIf(Len(rList) > 0, 1, 0)
                            packet(used, 15) = label: packet(used, 16) = start + r - 1
                            If used = 512 Then WriteMonitorPacket detail, packet, used, dest
                        End If
                    End If
                Next group
            Next r
            Application.StatusBar = "Match " & wt & " Function / Monitor Address: " & start + take - 1 & "/" & last
        Next start
    Next phase
    WriteMonitorPacket detail, packet, used, dest
    summaryDest = 5: summaryUsed = 0: ReDim sums(1 To 512, 1 To 16)
    For Each colKey In catalog
        entry = catalog(colKey)
        If CLng(entry(10)) <> CLng(entry(7)) Then Err.Raise vbObjectError + 495, , "Pattern Address count changed since coverage: " & CStr(colKey) & "; expected " & entry(7) & ", found " & entry(10) & ". Refresh analysis."
        summaryUsed = summaryUsed + 1
        sums(summaryUsed, 1) = entry(0): sums(summaryUsed, 2) = entry(1): sums(summaryUsed, 3) = entry(2)
        sums(summaryUsed, 4) = entry(3): sums(summaryUsed, 5) = entry(4)
        chipKey = CStr(entry(0)) & "|" & CStr(entry(1)) & "|" & CStr(entry(3)) & "|" & CStr(entry(4))
        sums(summaryUsed, 6) = chipCounts(chipKey): sums(summaryUsed, 7) = entry(5): sums(summaryUsed, 8) = entry(6)
        sums(summaryUsed, 9) = entry(10)
        If CLng(entry(10)) > 0 Then
            For j = 11 To 14: sums(summaryUsed, j - 1) = entry(j): Next j
            sums(summaryUsed, 16) = "Recorded Address classified"
        Else
            sums(summaryUsed, 16) = "No Function Address"
        End If
        sums(summaryUsed, 14) = entry(9): sums(summaryUsed, 15) = entry(8)
        If summaryUsed = 512 Then WriteMonitorPacket summary, sums, summaryUsed, summaryDest
    Next colKey
    WriteMonitorPacket summary, sums, summaryUsed, summaryDest
    For Each item In Array("Monitor Fail Analysis", "Monitor BLK Summary")
        With outputWorkbook.Worksheets(CStr(item))
            .Range("A1:P1").Merge: .Range("A1").Value2 = CStr(item)
            .Range("A2:P2").Merge: .Range("A2").Value2 = "Same WT + full X/Y/BLK/WL/BL/IO; any configured W-Low pBL / R-Low pBL voltage. No monitor item Fail = no recorded matching Monitor flag; unrecorded Function bits are not classified."
            .Range("A3:P3").Merge: .Range("A3").Value2 = "Map-counted BLKs only; Function Pattern interval union; exact numeric TNO mapping. Source: " & wb.Name
        End With
    Next item
    FormatMonitorOutput detail, dest - 1: FormatMonitorOutput summary, summaryDest - 1
    If dest > 5 Then
        With detail.Range("O5:O" & dest - 1)
            .FormatConditions.Add Type:=xlExpression, Formula1:="=$O5=""Write&Read Fail"""
            .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 237, 213)
        End With
    End If
    BLKVminEngine.BLK_OrderSheets outputWorkbook
    BLKVminEngine.BLK_BuildChipMonitorMap wb, outputWorkbook
    BuildMonitorFailAnalysis = "Monitor analysis complete. Map BLKs=" & catalog.Count & "; classified Addresses=" & total
    Exit Function
Failed:
    errorText = "Monitor analysis failed (" & Err.Number & "): " & Err.Description
    On Error Resume Next
    If Not detail Is Nothing Then detail.UsedRange.ClearContents: detail.Range("A2").Value2 = errorText
    If Not summary Is Nothing Then summary.UsedRange.ClearContents: summary.Range("A2").Value2 = errorText
    If Not outputWorkbook Is Nothing Then BLKVminEngine.BLK_ClearChipMonitorMap outputWorkbook
    BuildMonitorFailAnalysis = errorText
End Function

Public Sub WriteOperationLog(ByVal wb As Workbook, ByVal operation As String, ByVal message As String, Optional ByVal failed As Boolean = False)
    'One latest message per operation: bounded size, no per-Address log output.
    Dim ws As Worksheet, row As Long, values(1 To 1, 1 To 4) As Variant
    On Error Resume Next
    Set ws = wb.Worksheets("Log")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "Log"
    End If
    Select Case operation
        Case "Map": row = 2
        Case "FBC": row = 3
        Case "Address": row = 4
        Case Else: row = 5
    End Select
    ws.Range("A1:D1").Value2 = Array("Time", "Operation", "Status", "Message")
    values(1, 1) = Now: values(1, 2) = operation
    values(1, 3) = IIf(failed, "Warning / Error", "Complete")
    values(1, 4) = Left$(message, 32767)
    With ws.Range("A1:D5")
        .Font.Name = "Calibri": .Font.Size = 11: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ws.Cells(row, 1).Resize(1, 4).Value2 = values
    ws.Range("A2:A5").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Columns("A").ColumnWidth = 24: ws.Columns("B:C").ColumnWidth = 22: ws.Columns("D").ColumnWidth = 110
    ws.Range("D2:D5").WrapText = True: ws.Rows("2:5").RowHeight = 100: ws.Rows(1).RowHeight = 30
    With ws.Range("A1:D1")
        .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: .Borders.LineStyle = xlContinuous
    End With
    ws.Range("A1:D5").BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(80, 80, 80)
    If failed Then ws.Cells(row, 3).Interior.Color = RGB(255, 217, 102) Else ws.Cells(row, 3).Interior.Color = RGB(220, 252, 231)
    BLKVminEngine.BLK_SizeOutput ws
    BLKVminEngine.BLK_BuildImportInfo wb
End Sub

Public Function IsMonitorCountCell(ByVal ws As Object, ByVal target As Range) As Boolean
    Dim header As Range, wt As String
    On Error GoTo Done
    If TypeName(ws) <> "Worksheet" Or target.CountLarge <> 1 Then Exit Function
    If ws.Name <> "W&R Map" Or target.Row < 5 Then Exit Function
    If IsError(target.Value2) Then Exit Function
    If Not IsNumeric(target.Value2) Then Exit Function
    If CDbl(target.Value2) <= 0 Then Exit Function
    Set header = ws.Cells(3, target.Column).MergeArea.Cells(1, 1): wt = CStr(header.Value2)
    IsMonitorCountCell = (wt = "WT1" Or wt = "WT2" Or wt = "WT3") And target.Column >= header.Column
Done:
End Function

Public Sub OpenMonitorCountDetails(ByVal wb As Workbook, ByVal ws As Object, ByVal target As Range)
    Dim view As Worksheet, header As Range, wt As String, label As String, x As Long, y As Long, origin As Long, parity As Long
    On Error GoTo Failed
    If Not IsMonitorCountCell(ws, target) Then Exit Sub
    Set header = ws.Cells(3, target.Column).MergeArea.Cells(1, 1): wt = CStr(header.Value2): origin = header.Column - 1
    x = CLng(ws.Cells(4, target.Column).MergeArea.Cells(1, 1).Value2)
    y = CLng(ws.Cells(5 + 2 * ((target.Row - 5) \ 2), origin).MergeArea.Cells(1, 1).Value2)
    parity = (target.Column - header.Column) Mod 2
    If (target.Row - 5) Mod 2 = 0 Then
        If parity = 0 Then label = "Read Fail" Else label = "Write Fail"
    Else
        If parity = 0 Then label = "Write&Read Fail" Else label = "No monitor item Fail"
    End If
    Set view = wb.Worksheets("Monitor Fail Analysis")
    If Left$(CStr(view.Range("A2").Value2), 14) = "Source updated" Then Err.Raise vbObjectError + 498, , "Refresh analysis before opening Monitor detail."
    If view.FilterMode Then view.ShowAllData
    With view.Range("A4:P" & view.Cells(view.Rows.Count, 1).End(xlUp).Row)
        .AutoFilter Field:=1, Criteria1:=wt: .AutoFilter Field:=4, Criteria1:=x: .AutoFilter Field:=5, Criteria1:=y
        .AutoFilter Field:=15, Criteria1:=label
    End With
    view.Activate: Application.Goto view.Range("A4"), True
    Exit Sub
Failed:
    MsgBox "Cannot open Monitor detail: " & Err.Description, vbExclamation
End Sub

Public Function IsMapFailCell(ByVal mapSheet As Object, ByVal target As Range) As Boolean
    Dim groupHeader As Range, wt As String
    On Error GoTo InvalidCell
    If TypeName(mapSheet) <> "Worksheet" Or target.CountLarge <> 1 Then Exit Function
    If target.Row < 5 Or IsError(target.Value2) Then Exit Function
    If Not IsNumeric(target.Value2) Then Exit Function
    If CDbl(target.Value2) <= 0 Or CDbl(target.Value2) > 256 Then Exit Function
    If CDbl(target.Value2) <> Fix(CDbl(target.Value2)) Then Exit Function
    If DrillMapProperty(mapSheet, "ASCFailMap") <> "1" Then Exit Function
    Set groupHeader = mapSheet.Cells(3, target.Column).MergeArea.Cells(1, 1)
    wt = CStr(groupHeader.Value2)
    If wt <> "WT1" And wt <> "WT2" And wt <> "WT3" Then Exit Function
    If groupHeader.Column <= 1 Or target.Column < groupHeader.Column Then Exit Function
    IsMapFailCell = True
InvalidCell:
End Function

Public Function OpenMapChipDetails(ByVal wb As Workbook, ByVal mapSheet As Object, ByVal target As Range) As Boolean
    Dim wt As String, pattern As String, stopTno As String, dut As String, mask As String, key As Variant, tno As Variant
    Dim groupHeader As Range, masks As Object, wanted As Object, catalog As Object, addresses As Object, records As Object
    Dim stageSheet As Worksheet, oldSheet As Worksheet, source As Worksheet, cfg As Worksheet
    Dim x As Long, y As Long, origin As Long, start As Long, take As Long, last As Long, r As Long, b As Long
    Dim rowIndex As Long, detailHeader As Long, detailStart As Long, detailRow As Long, totalAddresses As Long, scopeFirst As Long, scopeLast As Long
    Dim data As Variant, info As Variant, entry As Variant, addressRow As Variant, ordered As New Collection, found As Object
    Dim packet() As Variant, detailPacket() As Variant, addressGroup As Collection, counts As Object, fbc As Double, gap As Double, finding As String
    Dim addressAvailable As Boolean, addressNote As String, oldEvents As Boolean, oldScreen As Boolean, oldAlerts As Boolean, ownsBusy As Boolean, published As Boolean
    Dim errorNumber As Long, errorText As String, savedTitle As String
    On Error GoTo Failed
    If TypeName(mapSheet) <> "Worksheet" Or target.CountLarge <> 1 Then Exit Function
    If target.Row < 5 Or IsError(target.Value2) Then Exit Function
    If Not IsNumeric(target.Value2) Then Exit Function
    If CDbl(target.Value2) <= 0 Then Exit Function
    If DrillMapProperty(mapSheet, "ASCFailMap") <> "1" Then Exit Function
    Set groupHeader = mapSheet.Cells(3, target.Column).MergeArea.Cells(1, 1)
    wt = CStr(groupHeader.Value2)
    If wt <> "WT1" And wt <> "WT2" And wt <> "WT3" Then Exit Function
    origin = groupHeader.Column - 1
    If origin < 1 Or target.Column <= origin Then Exit Function
    OpenMapChipDetails = True
    If busy Then Exit Function
    busy = True: ownsBusy = True
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating: oldAlerts = Application.DisplayAlerts
    Application.EnableEvents = False: Application.ScreenUpdating = False
    x = AuditInteger(mapSheet.Cells(4, target.Column).Value2, -2147483648#, 2147483647#)
    y = AuditInteger(mapSheet.Cells(target.Row, origin).Value2, -2147483648#, 2147483647#)
    pattern = DrillMapProperty(mapSheet, "ASCPattern"): stopTno = DrillMapProperty(mapSheet, "ASCStopTNO")
    If Len(pattern) = 0 Or Len(stopTno) = 0 Then Err.Raise vbObjectError + 500, , "Rebuild Maps once with this version to enable Chip details."
    Set cfg = wb.Worksheets("Config")
    Set wanted = DrillFunctionScope(cfg, stopTno, pattern, scopeFirst, scopeLast)
    dut = x & "|" & y: Set masks = BLKVminEngine.BLK_MapIncludedMasks(wb)
    savedTitle = wt & "|" & stopTno & "|" & dut
    If Not masks.Exists(savedTitle) Then Err.Raise vbObjectError + 500, , "This Map is unavailable with the current inputs. Rebuild Maps."
    mask = CStr(masks(savedTitle))
    If Len(Replace(mask, "0", "")) <> CLng(target.Value2) Then Err.Raise vbObjectError + 500, , "Map count changed. Rebuild Maps before opening details."
    Set catalog = CreateObject("Scripting.Dictionary"): Set found = CreateObject("Scripting.Dictionary")
    Set source = wb.Worksheets("FBC by BLK"): last = AuditLastRow(source, 2)
    For start = 2 To last Step 512
        take = last - start + 1: If take > 512 Then take = 512
        data = source.Cells(start, 1).Resize(take, 260).Value2
        For r = 1 To take
            tno = CStr(data(r, 2))
            If CStr(data(r, 1)) = wt And wanted.Exists(tno) Then
                If AuditInteger(data(r, 3), -2147483648#, 2147483647#) = x And AuditInteger(data(r, 4), -2147483648#, 2147483647#) = y Then
                    If found.Exists(tno) Then Err.Raise vbObjectError + 500, , "Duplicate Function WT/TNO/Chip FBC record."
                    found.Add tno, True: info = wanted(tno)
                    For b = 0 To 255
                        If Mid$(mask, b + 1, 1) = "1" Then
                            fbc = AuditFBCValue(data(r, b + 5))
                            If fbc > 0 Then catalog.Add tno & "|" & b, Array(CLng(tno), info(0), info(1), b, fbc)
                        End If
                    Next b
                End If
            End If
        Next r
    Next start
    For Each tno In wanted
        For b = 0 To 255
            If catalog.Exists(tno & "|" & b) Then ordered.Add tno & "|" & b
        Next b
    Next tno
    If ordered.Count = 0 Then Err.Raise vbObjectError + 500, , "No Function FBC details match this Chip. Rebuild Maps."
    Set addresses = CreateObject("Scripting.Dictionary"): Set counts = CreateObject("Scripting.Dictionary")
    addressNote = DrillLoadAddresses(wb, wanted, catalog, wt, x, y, addresses, counts, addressAvailable)
    For Each key In ordered
        If counts.Exists(key) Then totalAddresses = totalAddresses + CLng(counts(key))
    Next key
    detailHeader = 8 + ordered.Count: detailStart = detailHeader + 1
    If CDbl(detailStart) + totalAddresses > source.Rows.Count Then Err.Raise vbObjectError + 500, , "Chip Address details exceed Excel row limit."
    On Error Resume Next
    Set oldSheet = wb.Worksheets("Chip Fail Details")
    On Error GoTo Failed
    If Not oldSheet Is Nothing Then
        If DrillMapProperty(oldSheet, "ASCChipDetails") <> "1" Then Err.Raise vbObjectError + 500, , "Chip Fail Details sheet name is already used by another sheet."
    End If
    Set stageSheet = wb.Worksheets.Add(Before:=wb.Worksheets("Chip Locations")): stageSheet.Visible = xlSheetVeryHidden
    stageSheet.CustomProperties.Add Name:="ASCChipDetails", Value:="1"
    stageSheet.Range("A1:M1").Merge
    stageSheet.Range("A1").Value2 = pattern & " | " & wt & " | X=" & x & " Y=" & y & " | Map Fail BLKs=" & CLng(target.Value2)
    stageSheet.Range("A2:C2").Merge
    stageSheet.Hyperlinks.Add Anchor:=stageSheet.Range("A2"), Address:="", SubAddress:="'" & Replace(mapSheet.Name, "'", "''") & "'!" & target.Address(False, False), TextToDisplay:="Back to Map"
    stageSheet.Range("D2:M2").Merge: stageSheet.Range("D2").Value2 = "Function scope: Config D" & scopeFirst & ":D" & scopeLast & "; FBC records=" & found.Count & "/" & wanted.Count
    stageSheet.Range("A3:M3").Merge: stageSheet.Range("A3").Value2 = addressNote
    stageSheet.Range("A5:M5").Value2 = Array("WT", "Pattern", "TNO", "Func.", "Test Voltage", "X", "Y", "BLK", "FBC", "Address Count", "Missing Count", "Finding", "Address")
    ReDim packet(1 To ordered.Count, 1 To 13)
    If totalAddresses > 0 Then ReDim detailPacket(1 To totalAddresses, 1 To 11)
    rowIndex = 0: detailRow = 0
    For Each key In ordered
        rowIndex = rowIndex + 1: entry = catalog(key)
        packet(rowIndex, 1) = wt: packet(rowIndex, 2) = pattern: packet(rowIndex, 3) = entry(0): packet(rowIndex, 4) = entry(1)
        packet(rowIndex, 5) = entry(2): packet(rowIndex, 6) = x: packet(rowIndex, 7) = y: packet(rowIndex, 8) = entry(3): packet(rowIndex, 9) = entry(4)
        If addressAvailable Then
            b = 0: If counts.Exists(key) Then b = CLng(counts(key))
            gap = CDbl(entry(4)) - b: If gap < 0 Then gap = 0
            If b = 0 Then
                finding = "No Address"
            ElseIf b < CDbl(entry(4)) Then
                finding = "Partial Address"
            ElseIf b = CDbl(entry(4)) Then
                finding = "Complete"
            Else
                finding = "Address Count > FBC"
            End If
            packet(rowIndex, 10) = b: packet(rowIndex, 11) = gap: packet(rowIndex, 12) = finding
            If b > 0 Then
                Set addressGroup = addresses(key)
                For Each addressRow In addressGroup
                    detailRow = detailRow + 1
                    detailPacket(detailRow, 1) = wt: detailPacket(detailRow, 2) = entry(0): detailPacket(detailRow, 3) = entry(1): detailPacket(detailRow, 4) = entry(2)
                    detailPacket(detailRow, 5) = x: detailPacket(detailRow, 6) = y: detailPacket(detailRow, 7) = entry(3)
                    detailPacket(detailRow, 8) = addressRow(0): detailPacket(detailRow, 9) = addressRow(1): detailPacket(detailRow, 10) = addressRow(2): detailPacket(detailRow, 11) = addressRow(3)
                Next addressRow
            End If
        Else
            packet(rowIndex, 12) = "Address unavailable"
        End If
    Next key
    stageSheet.Range("A6").Resize(ordered.Count, 13).Value2 = packet
    'Write catalog links after Value2, which replaces their displayed text.
    rowIndex = 0: detailRow = 0
    For Each key In ordered
        rowIndex = rowIndex + 1
        If counts.Exists(key) Then
            b = CLng(counts(key))
            If b > 0 Then
                stageSheet.Hyperlinks.Add Anchor:=stageSheet.Cells(rowIndex + 5, 13), Address:="", SubAddress:="'Chip Fail Details'!A" & detailStart + detailRow, TextToDisplay:="Show Address"
                detailRow = detailRow + b
            End If
        End If
    Next key
    stageSheet.Cells(detailHeader - 1, 1).Resize(1, 11).Merge
    stageSheet.Cells(detailHeader - 1, 1).Value2 = "Detailed Fail Address | rows=" & totalAddresses & " | exact WT/TNO/Chip/BLK"
    stageSheet.Cells(detailHeader, 1).Resize(1, 11).Value2 = Array("WT", "TNO", "Func.", "Test Voltage", "X", "Y", "BLK", "WL", "BL", "IO", "Source Row")
    If totalAddresses > 0 Then
        stageSheet.Cells(detailStart, 1).Resize(totalAddresses, 11).Value2 = detailPacket
        For detailRow = 1 To totalAddresses
            stageSheet.Hyperlinks.Add Anchor:=stageSheet.Cells(detailStart + detailRow - 1, 11), Address:="", SubAddress:="'Fail Address'!A" & CLng(detailPacket(detailRow, 11)), TextToDisplay:=CStr(detailPacket(detailRow, 11))
        Next detailRow
    End If
    With stageSheet.UsedRange
        .Font.Name = "Calibri": .Font.Size = 11: .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 26
        .Columns.AutoFit
    End With
    For b = 1 To 13
        If stageSheet.Columns(b).ColumnWidth < 10 Then stageSheet.Columns(b).ColumnWidth = 10
        If stageSheet.Columns(b).ColumnWidth > 32 Then stageSheet.Columns(b).ColumnWidth = 32
    Next b
    stageSheet.Columns("D").ColumnWidth = 24: stageSheet.Columns("L").ColumnWidth = 26: stageSheet.Columns("M").ColumnWidth = 18
    stageSheet.Range("E6:E" & ordered.Count + 5).NumberFormat = "0.00#"
    If totalAddresses > 0 Then stageSheet.Cells(detailStart, 4).Resize(totalAddresses, 1).NumberFormat = "0.00#"
    stageSheet.Rows(1).RowHeight = 32: stageSheet.Rows(3).RowHeight = 42: stageSheet.Range("A3:M3").WrapText = True
    With stageSheet.Range("A1:M1"): .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: End With
    With stageSheet.Range("A5:M5"): .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: .Borders.LineStyle = xlContinuous: End With
    With stageSheet.Cells(detailHeader, 1).Resize(1, 11): .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: .Borders.LineStyle = xlContinuous: End With
    stageSheet.Range("A5:M" & ordered.Count + 5).AutoFilter
    stageSheet.UsedRange.BorderAround LineStyle:=xlContinuous, Weight:=xlThin, Color:=RGB(80, 80, 80)
    Application.DisplayAlerts = False
    If Not oldSheet Is Nothing Then oldSheet.Delete
    stageSheet.Name = "Chip Fail Details": stageSheet.Visible = xlSheetVisible: published = True
    BLKVminEngine.BLK_OrderSheets wb
    stageSheet.Activate
    With Application.ActiveWindow: .DisplayGridlines = False: .FreezePanes = False: .SplitColumn = 0: .SplitRow = 3: .FreezePanes = True: End With
    stageSheet.Range("A1").Select
    GoTo Cleanup
Failed:
    errorNumber = Err.Number: errorText = Err.Description
    Resume Cleanup
Cleanup:
    On Error Resume Next
    If ownsBusy Then
        Application.DisplayAlerts = False
        If Not stageSheet Is Nothing And Not published Then stageSheet.Visible = xlSheetVisible: stageSheet.Delete
        Application.DisplayAlerts = oldAlerts: Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen: busy = False
    End If
    If errorNumber <> 0 Then MsgBox "Chip details failed (" & errorNumber & "): " & errorText, vbExclamation
End Function

Private Function DrillMapProperty(ByVal ws As Object, ByVal propertyName As String) As String
    Dim i As Long
    For i = 1 To ws.CustomProperties.Count
        If ws.CustomProperties(i).Name = propertyName Then DrillMapProperty = CStr(ws.CustomProperties(i).Value): Exit Function
    Next i
End Function

Private Function DrillFunctionScope(ByVal cfg As Worksheet, ByVal stopTno As String, ByVal pattern As String, ByRef firstRow As Long, ByRef lastRow As Long) As Object
    Dim result As Object, data As Variant, i As Long, previous As Long, lastIndex As Long, tno As String
    data = cfg.Range("D2:G901").Value2: Set result = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(data, 1)
        If IsError(data(i, 1)) Or IsError(data(i, 4)) Then Err.Raise vbObjectError + 501, , "Invalid Function configuration."
        If Trim$(CStr(data(i, 1))) = stopTno And StrComp(Trim$(CStr(data(i, 4))), pattern, vbTextCompare) = 0 Then lastIndex = i: Exit For
        If Len(Trim$(CStr(data(i, 4)))) > 0 Then previous = i
    Next i
    If lastIndex = 0 Then Err.Raise vbObjectError + 501, , "Pattern configuration changed. Rebuild Maps."
    firstRow = previous + 2: lastRow = lastIndex + 1
    For i = previous + 1 To lastIndex
        tno = Trim$(CStr(data(i, 1)))
        If Len(tno) > 0 Then
            tno = CStr(AuditInteger(tno, 100, 999))
            If result.Exists(tno) Then Err.Raise vbObjectError + 501, , "Duplicate Function TNO."
            result.Add tno, Array(data(i, 2), data(i, 3))
        End If
    Next i
    Set DrillFunctionScope = result
End Function

Private Function DrillLoadAddresses(ByVal wb As Workbook, ByVal wanted As Object, ByVal catalog As Object, ByVal wt As String, ByVal x As Long, ByVal y As Long, ByVal addresses As Object, ByVal counts As Object, ByRef available As Boolean) As String
    Dim columns As Object, selected As Object, ws As Worksheet, groups As New Collection, group As Variant, info As Variant
    Dim lastCol As Long, lastRow As Long, col As Long, firstCol As Long, batch As Long, maximumWidth As Long, start As Long, take As Long, r As Long, j As Long, b As Long
    Dim coords As Variant, flags As Variant, flag As Variant, candidate() As Boolean, hasCandidate As Boolean, key As String, note As String, sourceNote As String
    Dim rows As Collection, errorText As String
    On Error GoTo Failed
    sourceNote = AuditSourceNote(wb): Set columns = ResolveAuditColumns(wb, wanted, note)
    Set selected = CreateObject("Scripting.Dictionary")
    For Each group In columns
        info = columns(group)
        If CStr(info(0)) = wt Then selected.Add CStr(group), info
    Next group
    Set ws = wb.Worksheets("Fail Address"): lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    col = 7
    Do While col <= lastCol
        If selected.Exists(CStr(col)) Then
            firstCol = col
            Do While col <= lastCol
                If Not selected.Exists(CStr(col)) Then Exit Do
                col = col + 1
            Loop
            groups.Add Array(firstCol, col - firstCol)
            If col - firstCol > maximumWidth Then maximumWidth = col - firstCol
        Else
            col = col + 1
        End If
    Loop
    lastRow = AuditLastRow(ws, 1)
    batch = 2048: If maximumWidth > 0 Then If batch * maximumWidth > 131072 Then batch = 131072 \ maximumWidth
    For start = 2 To lastRow Step batch
        take = lastRow - start + 1: If take > batch Then take = batch
        coords = ws.Cells(start, 1).Resize(take, 6).Value2: ReDim candidate(1 To take): hasCandidate = False
        For r = 1 To take
            If AuditInteger(coords(r, 1), -2147483648#, 2147483647#) = x And AuditInteger(coords(r, 2), -2147483648#, 2147483647#) = y Then
                candidate(r) = True: hasCandidate = True
            End If
        Next r
        If hasCandidate Then
            For Each group In groups
                flags = ws.Cells(start, CLng(group(0))).Resize(take, CLng(group(1))).Value2
                For j = 1 To CLng(group(1))
                    info = selected(CStr(CLng(group(0)) + j - 1))
                    For r = 1 To take
                        If candidate(r) Then
                            b = AuditInteger(coords(r, 3), 0, 255): key = CStr(info(1)) & "|" & b
                            If catalog.Exists(key) Then
                                If take = 1 And CLng(group(1)) = 1 Then flag = flags Else flag = flags(r, j)
                                If IsError(flag) Then Err.Raise vbObjectError + 502, , "Address flag error."
                                If Len(CStr(flag)) > 0 Then
                                    If Not IsNumeric(flag) Then Err.Raise vbObjectError + 502, , "Address flag must be0/1."
                                    If CDbl(flag) <> 0 And CDbl(flag) <> 1 Then Err.Raise vbObjectError + 502, , "Address flag must be0/1."
                                    If CDbl(flag) = 1 Then
                                        If Not counts.Exists(key) Then
                                            counts.Add key, 0: Set rows = New Collection: addresses.Add key, rows
                                        End If
                                        counts(key) = CLng(counts(key)) + 1: Set rows = addresses(key)
                                        rows.Add Array(AuditInteger(coords(r, 4), 0, 1023), AuditInteger(coords(r, 5), 0, 15), AuditInteger(coords(r, 6), 0, 15), start + r - 1)
                                    End If
                                End If
                            End If
                        End If
                    Next r
                Next j
            Next group
        End If
    Next start
    available = True: DrillLoadAddresses = note & " " & sourceNote
    Exit Function
Failed:
    errorText = Err.Description: counts.RemoveAll: addresses.RemoveAll: available = False
    DrillLoadAddresses = "Address unavailable: " & errorText & " FBC details remain available."
End Function


Private Sub RecordImportIndex(ByVal wb As Workbook, ByVal mode As String, ByRef meta() As Object, ByRef paths() As String, ByVal tnos As Object, Optional ByVal addressPhases As Long = 0)
    Dim ws As Worksheet, mapping() As Variant, manifest(1 To 3, 1 To 8) As Variant, headers As Variant, phase As Long, tno As Variant, i As Long, col As Long, start As Long
    Set ws = wb.Worksheets("Import Index")
    If addressPhases = 0 Then addressPhases = BLKVminEngine.BLK_PhaseCount(wb)
    If mode = "Address" Then
        ws.Range("A:D").ClearContents
        ws.Range("A1:D1").Value2 = Array("WT", "TNO", "Column", "Header")
        ReDim mapping(1 To tnos.Count * addressPhases, 1 To 4)
        headers = wb.Worksheets("Fail Address").Cells(1, 1).Resize(1, 7 + tnos.Count * addressPhases).Value2
        For phase = 1 To addressPhases
            For Each tno In tnos
                i = i + 1: col = 8 + (phase - 1) * tnos.Count + CLng(tnos(tno))
                mapping(i, 1) = "WT" & phase: mapping(i, 2) = CStr(tno): mapping(i, 3) = col: mapping(i, 4) = headers(1, col)
            Next tno
        Next phase
        ws.Range("A2").Resize(i, 4).Value2 = mapping
    End If
    For phase = 1 To 3
        manifest(phase, 1) = mode: manifest(phase, 2) = "WT" & phase
        If meta(phase) Is Nothing Then
            manifest(phase, 1) = mode & " Not tested"
        Else
            manifest(phase, 3) = CStr(meta(phase)("LOT"))
            manifest(phase, 4) = NormalWafer(CStr(meta(phase)("WNO"))): manifest(phase, 5) = CStr(meta(phase)("STEP"))
            manifest(phase, 6) = BuildResultBaseName(meta(phase)): manifest(phase, 7) = paths(phase): manifest(phase, 8) = Now
        End If
    Next phase
    ws.Range("H1:O1").Value2 = Array("Mode", "WT", "Full LOT", "WNO", "STEP", "Identity", "Source path", "Imported at")
    start = IIf(mode = "Address", 2, 5)
    ws.Cells(start, 8).Resize(3, 8).Value2 = manifest
    ws.Visible = xlSheetVeryHidden
    If mode = "Address" Then AddFailAddressNotes wb
    MarkAuditStale wb, mode
    BLKVminEngine.BLK_InvalidateDetails wb, mode & " source updated. Double-click a Fail Chip to refresh details."
End Sub

Private Sub AddFailAddressNotes(ByVal wb As Workbook)
    Dim idx As Worksheet, ws As Worksheet, data As Variant, i As Long, last As Long, col As Long, monitors As Object, cfg As Variant
    Set idx = wb.Worksheets("Import Index"): Set ws = wb.Worksheets("Fail Address")
    ws.Rows(1).ClearComments: last = idx.Cells(idx.Rows.Count, 1).End(xlUp).Row
    If last < 2 Then Exit Sub
    Set monitors = CreateObject("Scripting.Dictionary"): cfg = wb.Worksheets("Config").Range("A2:A901").Value2
    For i = 1 To UBound(cfg, 1): If Len(CStr(cfg(i, 1))) > 0 Then monitors(CStr(CLng(cfg(i, 1)))) = True
    Next i
    data = idx.Range("A2:D" & last).Value2
    For i = 1 To UBound(data, 1)
        col = CLng(data(i, 3))
        ws.Cells(1, col).AddComment CStr(data(i, 1)) & " | " & IIf(monitors.Exists(CStr(data(i, 2))), "Monitor", "Function") & " | TNO=" & CStr(data(i, 2))
    Next i
End Sub

Private Sub MarkAuditStale(ByVal wb As Workbook, ByVal mode As String)
    On Error Resume Next
    With wb.Worksheets("Wafer Maps")
        .Range("A2").Value2 = mode & " source updated. Refresh Maps before using these figures."
        .Range("A2:U2").Interior.Color = RGB(255, 217, 102)
    End With
    With wb.Worksheets("Fail Analysis")
        .Range("A3").Value2 = mode & " source updated. Refresh Map / Address comparison before using these rows."
        .Range("A3:P3").Interior.Color = RGB(255, 217, 102)
    End With
    With wb.Worksheets("Missing BLK Address")
        .Range("A2").Value2 = mode & " source import updated. Refresh BLK Address comparison before using its rows."
        .Range("A2:L2").Interior.Color = RGB(255, 217, 102)
    End With
    With wb.Worksheets("FailBLK Loss Map")
        If Len(CStr(.Range("A1").Value2)) > 0 Then
            .Range("A1").Value2 = "Address/FBC updated - refresh Map"
            .Range("A1").Interior.Color = RGB(255, 217, 102): .Range("A1").Font.Color = vbBlack
        End If
    End With
    With wb.Worksheets("Monitor Fail Analysis")
        .Range("A2").Value2 = "Source updated. Refresh Map / Address comparison to update Monitor analysis."
        .Range("A2:P2").Interior.Color = RGB(255, 217, 102)
    End With
    With wb.Worksheets("Monitor BLK Summary")
        .Range("A2").Value2 = "Source updated. Refresh Map / Address comparison to update Monitor analysis."
        .Range("A2:P2").Interior.Color = RGB(255, 217, 102)
    End With

    With wb.Worksheets("W&R Map")
        If Len(CStr(.Range("A1").Value2)) > 0 Then
            .Range("A1").Value2 = "Monitor source updated - refresh Map"
            .Range("A1").Interior.Color = RGB(255, 217, 102): .Range("A1").Font.Color = vbBlack
        End If
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

Public Sub GoWaferMap()
    'Map navigation is an explicit user action. Recover the event switch if a previous macro left it off.
    If VarType(Application.Caller) = vbString Then Application.EnableEvents = True
    BLKVminEngine.BLK_GoWaferMap ThisWorkbook
End Sub
Public Sub OpenSelectedChipDetails()
    Dim ws As Worksheet, selectedCell As Range
    If TypeName(Application.ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = Application.ActiveSheet
    If Not ws.Parent Is ThisWorkbook Then Exit Sub
    If ws.Name <> "Wafer Maps" Then Exit Sub
    Set selectedCell = Application.ActiveCell
    If Not BLKVminEngine.BLK_IsIndexedMapCell(ThisWorkbook, ws, selectedCell) Then
        MsgBox "Select a Chip cell on the Map, then click Chip Details.", vbInformation
        Exit Sub
    End If
    BLKVminEngine.BLK_OpenIndexedMap ThisWorkbook, ws, selectedCell
End Sub

Public Sub ExportWaferMapsPPT()
    Dim path As String, started As Date
    On Error GoTo Failed
    started = Now
    path = BLKVminEngine.BLK_ExportMapsPPT(ThisWorkbook)
    MsgBox "PPT saved: " & path & vbCrLf & "Seconds: " & DateDiff("s", started, Now), vbInformation
    Exit Sub
Failed:
    MsgBox "PPT export failed: " & Err.Description, vbExclamation
End Sub
