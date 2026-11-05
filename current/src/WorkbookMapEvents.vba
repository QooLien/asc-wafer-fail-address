Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim detailMessage As String
    If BLKVminEngine.BLK_IsVoltageGridCell(Sh, Target) Then
        Cancel = True
        detailMessage = BLKVminEngine.BLK_ShowBLKVoltage(Me, Sh, Target)
        If Left$(detailMessage, 6) = "Error:" Then MsgBox detailMessage, vbExclamation
    ElseIf BLKVminEngine.BLK_IsIndexedMapCell(Me, Sh, Target) Then
        Cancel = True
        BLKVminEngine.BLK_OpenIndexedMap Me, Sh, Target
    ElseIf ImportFailAddresses.IsMonitorCountCell(Sh, Target) Then
        Cancel = True
        ImportFailAddresses.OpenMonitorCountDetails Me, Sh, Target
    ElseIf ImportFailAddresses.IsMapFailCell(Sh, Target) Then
        Cancel = True
        ImportFailAddresses.OpenMapChipDetails Me, Sh, Target
    End If
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If BLKVminEngine.BLK_SourceRowCell(Sh, Target) Then BLKVminEngine.BLK_JumpSourceRow Me, Target
End Sub
