"""Create a separate classified variant without editing the verified importer."""
from pathlib import Path
root=Path(__file__).resolve().parents[1]
source=(root/'src/ImportThreeASC.bas').read_text()
source=source.replace('    Dim cfg As Worksheet', '''    Dim group As Long, groups() As Byte, groupCounts(1 To 3) As Long, judged As Long
    Dim conflictCount As Long, conflictExamples As String
    Dim cfg As Worksheet''')
source=source.replace('Set dst = ThisWorkbook.Worksheets("Fail Address")','Set dst = ThisWorkbook.Worksheets("WT1 Fail")')
source=source.replace(' & "_WT1-WT3"',' & "_WT_Classified"')
position=source.index('    columnCount = 6 + totalTNOs')
source=source[:position]+'''    stage = "Classify selected-TNO fail stages"
    If uniqueCount > 0 Then ReDim groups(0 To uniqueCount - 1)
    For Each v In addresses
        rowIndex = CLng(addresses(v))
        judged = FailWT(bits, rowIndex, tnos.Count)
        If judged <= 0 Then
            conflictCount = conflictCount + 1
            If conflictCount <= 5 Then conflictExamples = conflictExamples & vbCrLf & CStr(v)
        Else
            groups(rowIndex) = CByte(judged)
            groupCounts(judged) = groupCounts(judged) + 1
        End If
    Next v
    If conflictCount > 0 Then Err.Raise vbObjectError + 30, , "Classification conflict: " & conflictCount & " addresses have fails in multiple WTs (or no selected fail). Previous output retained. First addresses X,Y,BLK,WL,BL,IO:" & conflictExamples
''' +source[position:]
start=source.index('    stage = "Write output batches"')
end=source.index('    cfg.Range("D2")',start)
block=source[start:end]
block=block.replace('uniqueCount','groupCounts(group)')
block=block.replace('            batchRow = batchRow + 1','            If groups(rowIndex) = group Then\n            batchRow = batchRow + 1')
block=block.replace('        Next v','            End If\n        Next v')
source=source[:start]+'''    For group = 1 To 3
        Set dst = ThisWorkbook.Worksheets("WT" & group & " Fail")
        written = 0: batchRow = 0
'''+block+'''    Next group
    cfg.Range("G2:L4").Value2 = sourceStats
    cfg.Range("D7").Value2 = groupCounts(1)
    cfg.Range("D8").Value2 = groupCounts(2)
    cfg.Range("D9").Value2 = groupCounts(3)
''' +source[end:]
pos=source.index("' Only orchestration helpers")
helper='''
' Judge ONLY selected Config TNOs. Multiple fails in one WT are allowed.
' A fail in two or more WTs is a flow conflict, never silently assigned.
Private Function FailWT(ByRef bits() As Byte, ByVal rowIndex As Long, ByVal tnoCount As Long) As Long
    Dim phase As Long, b As Long, mask As Long, first As Long, last As Long, found As Boolean, result As Long
    For phase = 1 To 3
        first = (phase - 1) * tnoCount: last = first + tnoCount - 1
        found = False
        For b = first \\ 8 To last \\ 8
            mask = 255
            If b = first \\ 8 Then mask = mask And (256 - mBitMasks(first And 7))
            If b = last \\ 8 Then mask = mask And (mBitMasks(last And 7) * 2 - 1)
            If (bits(b, rowIndex) And mask) <> 0 Then
                found = True
                Exit For
            End If
        Next b
        If found Then
            If result <> 0 Then FailWT = -1: Exit Function
            result = phase
        End If
    Next phase
    FailWT = result
End Function

'''
source=source[:pos]+helper+source[pos:]
target=root/'src/ImportThreeASC_Classified.bas'
target.write_text(source,encoding='utf-8')
base=(root/'src/ImportThreeASC.bas').read_text()
assert source[source.index('Private Function PhaseTNOs'):]==base[base.index('Private Function PhaseTNOs'):]
print('PASS: classified importer created; original phase mapping, reader and parser helpers unchanged.')
