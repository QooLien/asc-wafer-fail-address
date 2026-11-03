Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    If ImportFailAddresses.IsMapFailCell(Sh, Target) Then
        Cancel = True
        ImportFailAddresses.OpenMapChipDetails Me, Sh, Target
    End If
End Sub
