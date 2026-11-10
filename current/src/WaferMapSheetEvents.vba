Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Cancel Then Exit Sub
    If Not Application.ActiveSheet Is Me Then Exit Sub
    If Not BLKVminEngine.BLK_IsIndexedMapCell(ThisWorkbook, Me, Target) Then Exit Sub
    Cancel = True
    BLKVminEngine.BLK_OpenIndexedMap ThisWorkbook, Me, Target
End Sub
