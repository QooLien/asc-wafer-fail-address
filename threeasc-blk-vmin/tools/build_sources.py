"""Integrate only import-three-ASC and per-WT BLK Vmin; do not include split modules."""
from pathlib import Path
root=Path(__file__).resolve().parents[1]
old=root/'reference/ImportThreeASC.bas'
blk=root/'reference/BLKDistribution.bas'
source=old.read_text()
source=source.replace('    Dim cfg As Worksheet', '    Dim blkWarnings As String, blkMessage As String\n    Dim cfg As Worksheet')
source=source.replace('    stage = "Select three ASC files"', '''    stage = "Validate BLK Vmin settings"
    BLKVminEngine.BLK_Setup cfg
    stage = "Select three ASC files"''')
source=source.replace(' & "_WT1-WT3"', ' & "_WT1-WT3_BLK_Vmin"')
source=source.replace('        DetectEncoding fileNo, encoding\n    Do While', '        DetectEncoding fileNo, encoding\n        BLKVminEngine.BLK_BeginFile path\n    Do While')
source=source.replace('        buffer = buffer & raw\n        raw = ""', '        buffer = buffer & raw\n        BLKVminEngine.BLK_Feed raw, False\n        raw = ""')
source=source.replace('        allAccepted = allAccepted + accepted: allSkipped = allSkipped + skipped', '''        stage = "BLK Vmin WT" & phase
        BLKVminEngine.BLK_Feed "", True
        blkMessage = BLKVminEngine.BLK_WritePhase(ThisWorkbook.Worksheets("WT" & phase & "_BLK_Distribution"))
        If Len(blkMessage) > 0 Then blkWarnings = blkWarnings & vbCrLf & "WT" & phase & ": " & blkMessage
        allAccepted = allAccepted + accepted: allSkipped = allSkipped + skipped''')
source=source.replace('    stage = "Save result workbook"', '    cfg.Range("Q2").Value2 = blkWarnings\n    stage = "Save result workbook"')
source=source.replace('& "No log output.", vbInformation', '& "No log output." & blkWarnings, vbInformation')
source=source.replace('    busy = False\n    Exit Sub', '    BLKVminEngine.BLK_Release\n    busy = False\n    Exit Sub')
(root/'src').mkdir(parents=True,exist_ok=True)
(root/'src/ImportThreeASC_BLKVmin.bas').write_text(source,encoding='utf-8')
base=blk.read_text()
declarations=base[:base.index('Public Sub ImportASCFailAddresses')].replace('Attribute VB_Name = "ImportFailAddresses"','Attribute VB_Name = "BLKVminEngine"')
declarations=declarations.replace('Option Explicit','Option Explicit\nOption Private Module')
declarations+='''Private mSequence() As String, mVoltage() As Double, mConfiguredDUTs As Collection
Private mThreshold As Double, mCarry As String, mBLKError As String

'''
config=base[base.index('Private Sub ReadConfiguration'):base.index('Private Function ProcessFile')]
config=config.replace('DataPath','Config').replace('ws.Range("D2")','ws.Range("N2")')
config=config.replace('ws.Rows.Count, 6','ws.Rows.Count, 14').replace('ws.Rows.Count, 7','ws.Rows.Count, 15').replace('"F2:G"','"N6:O"')
# DUT list begins row 6; account for list height and human-facing row numbers.
config=config.replace('If last >= 2 Then','If last >= 6 Then').replace('i + 1\n            If Len(Trim$(CStr(xy','i + 5\n            If Len(Trim$(CStr(xy')
config=config.replace('    data = ws.Range("A2:B"', '    If last > 901 Then Err.Raise vbObjectError + 319, , "Use Config!A2:B901 for both address and BLK TNOs."\n    data = ws.Range("A2:B"')
config=config.replace('If tno = "0" Then', 'If CLng(tno) < 100 Or CLng(tno) > 999 Then').replace('TNO must be a positive integer.','Configured TNO must be 100-999.')
parser=base[base.index('Private Sub ResetParser'):base.index('Private Sub WriteGroup')]
writer=base[base.index('Private Sub WriteGroup'):base.index('Private Function UniquePath')]
wrappers='''Public Sub BLK_Setup(ByVal cfg As Worksheet)
    ReadConfiguration cfg, mSequence, mVoltage, mConfiguredDUTs, mThreshold
End Sub

Public Sub BLK_BeginFile(ByVal path As String)
    Dim i As Long
    Set mWanted = CreateObject("Scripting.Dictionary")
    For i = LBound(mSequence) To UBound(mSequence): mWanted.Add mSequence(i), True: Next i
    ResetParser path
    mCarry = "": mBLKError = ""
End Sub

Public Sub BLK_Feed(ByVal raw As String, ByVal atEOF As Boolean)
    Dim p As Long, line As String
    If Len(mBLKError) > 0 Then Exit Sub
    On Error GoTo Failed
    mCarry = mCarry & raw
    Do
        p = InStr(mCarry, vbLf)
        If p = 0 Then Exit Do
        mLine = mLine + 1: line = Replace(Left$(mCarry, p - 1), vbCr, "")
        If mActive Or InStr(line, "=") > 0 Then ParseLine line
        mCarry = Mid$(mCarry, p + 1)
    Loop
    If Len(mCarry) > MAX_LINE Then Err.Raise vbObjectError + 310, , "ASC line exceeds 1 MB."
    If atEOF Then
        If Len(mCarry) > 0 Then
            mLine = mLine + 1: line = Replace(mCarry, vbCr, "")
            If mActive Or InStr(line, "=") > 0 Then ParseLine line
        End If
        FinishRecord: mCarry = ""
    End If
    Exit Sub
Failed:
    mBLKError = "FBC Parse Error (" & Err.Number & "): " & Err.Description & "; line=" & mLine
    mCarry = "": mActive = False
End Sub

Public Function BLK_WritePhase(ByVal ws As Worksheet) As String
    Dim duts As Collection, key As Variant, i As Long, block As Long, groupSize As Long, firstRow As Long
    Dim totalRows As Double, header(1 To 1, 1 To 260) As Variant, group As Variant, errorText As String
    On Error GoTo Failed
    ws.UsedRange.UnMerge: ws.UsedRange.Clear
    header(1, 1) = "TNO / Vmin Label": header(1, 2) = "Test Voltage": header(1, 3) = "DUT X": header(1, 4) = "DUT Y"
    For block = 0 To 255: header(1, block + 5) = "BLK " & block: Next block
    ws.Range("A1:IZ1").Value2 = header
    With ws.Range("A1:IZ2")
        .Font.Name = "Calibri": .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 26
    End With
    ws.Columns("A").ColumnWidth = 26: ws.Columns("B").ColumnWidth = 15
    ws.Columns("C:D").ColumnWidth = 10: ws.Columns("E:IZ").ColumnWidth = 11
    ws.Range("A1:IZ1").Interior.Color = RGB(23, 50, 77): ws.Range("A1:IZ1").Font.Color = vbWhite
    If Len(mBLKError) > 0 Then
        ws.Range("A2:D2").Merge: ws.Range("A2").Value2 = mBLKError
        ws.Range("A2:D2").WrapText = True: ws.Rows(2).RowHeight = 60
        BLK_WritePhase = mBLKError: Exit Function
    End If
    Set duts = New Collection
    If mConfiguredDUTs.Count > 0 Then
        For Each key In mConfiguredDUTs: duts.Add CStr(key): Next key
    Else
        Set duts = SortedDUTs(mDUTs)
    End If
    If duts.Count = 0 Then
        ws.Range("A2:D2").Merge: ws.Range("A2").Value2 = "No Data: no FBC BY BLK records for configured TNOs."
        ws.Range("A2:D2").WrapText = True: ws.Rows(2).RowHeight = 60
        BLK_WritePhase = "No selected FBC BY BLK data; Fail Address import preserved.": Exit Function
    End If
    groupSize = UBound(mSequence) + 1: totalRows = CDbl(duts.Count) * groupSize + 1
    If totalRows > ws.Rows.Count Then Err.Raise vbObjectError + 308, , "BLK output exceeds worksheet row limit."
    firstRow = 2
    For Each key In duts
        group = BuildDUTGroup(mData, mSequence, mVoltage, CStr(key), mThreshold)
        WriteGroup ws, firstRow, group: firstRow = firstRow + groupSize
    Next key
    With ws.Range("A1:IZ" & CLng(totalRows))
        .Font.Name = "Calibri": .Font.Size = 11
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .RowHeight = 26
    End With
    ws.Columns("A").ColumnWidth = 26: ws.Columns("B").ColumnWidth = 15
    ws.Columns("C:D").ColumnWidth = 10: ws.Columns("E:IZ").ColumnWidth = 11
    ws.Columns("B").NumberFormat = "0.000"
    With ws.Range("A1:IZ1")
        .Interior.Color = RGB(23, 50, 77): .Font.Color = vbWhite: .Font.Bold = True: .RowHeight = 32
    End With
    firstRow = 2 + groupSize - 1
    For i = 1 To duts.Count
        ws.Cells(firstRow, 1).Resize(1, 4).Merge
        With ws.Cells(firstRow, 1).Resize(1, 260)
            .Interior.Color = RGB(37, 99, 168): .Font.Color = vbWhite: .Font.Bold = True
        End With
        firstRow = firstRow + groupSize
    Next i
    Exit Function
Failed:
    errorText = "BLK Output Error (" & Err.Number & "): " & Err.Description
    On Error Resume Next
    ws.UsedRange.UnMerge: ws.UsedRange.ClearContents
    ws.Range("A1:IZ1").Value2 = header
    ws.Range("A2:D2").Merge: ws.Range("A2").Value2 = errorText
    BLK_WritePhase = errorText
End Function

Public Sub BLK_Release()
    Set mData = Nothing: Set mWanted = Nothing: Set mDUTs = Nothing
    mCarry = "": mActive = False
End Sub

'''
(root/'src/BLKVminEngine.bas').write_text(declarations+wrappers+config+parser+writer,encoding='utf-8')
assert source[source.index('Private Function PhaseTNOs'):]==old.read_text()[old.read_text().index('Private Function PhaseTNOs'):]
print('PASS: integrated sources generated; original three-ASC helpers unchanged; standalone BLK parser/Judge reused.')
