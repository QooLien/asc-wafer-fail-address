"""Derive an isolated three-ASC importer; keep every original helper byte-for-byte."""
from pathlib import Path
root=Path(__file__).resolve().parents[1]
base=(root/'src/ImportFailAddresses.bas').read_text(encoding='utf-8')
main=base[:base.index('Private Sub DetectEncoding')]
helpers=base[base.index('Private Sub DetectEncoding'):]
main=main.replace('    Dim cfg As Worksheet', '''    Dim paths(1 To 3) As String, meta(1 To 3) As Object, steps(1 To 3) As Long
    Dim phase As Long, other As Long, swapPath As String, swapStep As Long, swapMeta As Object
    Dim localTNOs As Object, allAccepted As Long, allSkipped As Long, sourceStats(1 To 3, 1 To 6) As Variant
    Dim baseNames(1 To 3) As String, totalTNOs As Long, color As Long
    Dim cfg As Worksheet''')
a=main.index('    stage = "Select ASC"')
b=main.index('    flagBytes =',a)
main=main[:a]+'''    stage = "Select three ASC files"
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
''' + main[b:]
main=main.replace('flagBytes = (tnos.Count + 7)', 'flagBytes = (totalTNOs + 7)')
a=main.index('    Do While Seek(fileNo)')
b=main.index('    uniqueCount = addresses.Count',a)
oldloop=main[a:main.index('    Close #fileNo',a)]
m1=oldloop.index('        If metadata.Count < 4 Then')
m2=oldloop.index('        buffer = buffer & raw',m1)
oldloop=oldloop[:m1]+oldloop[m2:]
oldloop=oldloop.replace('skipped, tnos, addresses','skipped, localTNOs, addresses')
main=main[:a]+'''    For phase = 1 To 3
        path = paths(phase): buffer = "": headerTNO = ""
        selectedSections = 0: accepted = 0: skipped = 0: bytePosition = 0
        Set localTNOs = PhaseTNOs(tnos, phase)
        stage = "Open WT" & phase
        fileNo = FreeFile: Open path For Binary Access Read As #fileNo
        fileOpen = True: fileSize = LOF(fileNo)
        DetectEncoding fileNo, encoding
'''+oldloop+'''        Close #fileNo: fileOpen = False
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
''' +main[b:]
main=main.replace('columnCount = 6 + tnos.Count','columnCount = 6 + totalTNOs')
main=main.replace('    For Each v In tnos: headers(1, 7 + CLng(tnos(v))) = CLng(v): Next v', '''    For phase = 1 To 3
        Set localTNOs = PhaseTNOs(tnos, phase)
        For Each v In localTNOs: headers(1, 7 + CLng(localTNOs(v))) = "WT" & phase & vbLf & CStr(v): Next v
    Next phase''')
main=main.replace('    dst.UsedRange.ClearContents','    dst.UsedRange.FormatConditions.Delete\n    dst.UsedRange.ClearContents')
a=main.index('    With dst.Cells(1, 7)')
b=main.index('    dst.Cells(1, 1).Resize(uniqueCount',a)
main=main[:a]+'''    For phase = 1 To 3
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
''' +main[b:]
main=main.replace('cfg.Range("D2").Value2 = path','cfg.Range("D2").Value2 = "3 ASC files; see WT source mapping"')
main=main.replace('cfg.Range("D5").Value2 = accepted','cfg.Range("D5").Value2 = allAccepted').replace('cfg.Range("D6").Value2 = skipped','cfg.Range("D6").Value2 = allSkipped')
extra='''
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

'''
target=root/'src/ImportThreeASC.bas'
target.write_text(main+extra+helpers,encoding='utf-8')
assert target.read_text()[target.read_text().index('Private Sub DetectEncoding'):] == helpers
print('PASS: generated isolated three-file VBA; original reader/metadata/parser helpers unchanged.')
