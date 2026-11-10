Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not BLKVminEngine.BLK_IsIndexedMapCell(ThisWorkbook, Me, Target) Then Exit Sub
    Cancel = True
    BLKVminEngine.BLK_OpenIndexedMap ThisWorkbook, Me, Target
End Sub
