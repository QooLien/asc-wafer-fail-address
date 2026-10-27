Attribute VB_Name = "ImportFailAddresses"
Option Explicit

' Windows desktop Excel. Import this module into the supplied workbook.
' Config!A2:A901: selected TNOs. Fail Addresses: one row per unique address.
Public Sub ImportASCFailAddresses()
    Dim cfg As Worksheet, dst As Worksheet, tnos As Object, addresses As Object
    Dim fd As Object, fso As Object, stream As Object, re As Object, tokens As Object
    Dim token As Object, flags As Object, fields As Variant, addr As Variant, keys As Variant
    Dim raw As String, path As String, s As String, key As String, tno As String
    Dim r As Long, j As Long, col As Long, headerTNO As String, n As Double
    Dim nums(0 To 7) As Long, valid As Boolean, v As Variant
    Dim skipped As Long, accepted As Long, result() As Variant, oldScreen As Boolean
    On Error GoTo Failed
    oldScreen = Application.ScreenUpdating
    Set cfg = ThisWorkbook.Worksheets("Config")
    Set dst = ThisWorkbook.Worksheets("Fail Addresses")
    Set tnos = CreateObject("Scripting.Dictionary")
    Set addresses = CreateObject("Scripting.Dictionary")
    For r = 2 To 901
        v = cfg.Cells(r, 1).Value2
        If IsError(v) Then Err.Raise vbObjectError + 1, , "Invalid TNO at Config!A" & r
        s = Trim$(CStr(v))
        If Len(s) > 0 Then
            If Not IntegerInRange(s, 100, 999, n) Or Len(s) <> 3 Then
                Err.Raise vbObjectError + 2, , "TNO must be 100-999 at Config!A" & r
            End If
            tno = CStr(CLng(n))
            If Not tnos.Exists(tno) Then tnos.Add tno, tnos.Count + 7
        End If
    Next r
    If tnos.Count = 0 Then Err.Raise vbObjectError + 3, , "Enter TNOs in Config!A2:A901."
    Set fd = Application.FileDialog(3)
    With fd
        .Title = "Select ASC raw data"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "ASC raw data", "*.asc"
        If .Show <> -1 Then Exit Sub
        path = .SelectedItems(1)
    End With
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set stream = fso.OpenTextFile(path, 1, False, -2)
    raw = stream.ReadAll
    stream.Close
    Set stream = Nothing
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True
    ' Ordered tokens preserve each tno= section, including multiline brackets.
    re.Pattern = "\btno\s*=\s*[0-9]+\b|\[[^\]]*\]"
    Set tokens = re.Execute(raw)
    For Each token In tokens
        s = CStr(token.Value)
        If Left$(s, 1) <> "[" Then
            headerTNO = Trim$(Replace(Replace(Replace(Mid$(s, InStr(s, "=") + 1), vbTab, " "), vbCr, " "), vbLf, " "))
        ElseIf tnos.Exists(headerTNO) Then
            fields = Split(Mid$(s, 2, Len(s) - 2), ",")
            ' Exactly 8 values, with an optional trailing comma.
            If UBound(fields) = 8 Then
                If Len(Trim$(CStr(fields(8)))) = 0 Then ReDim Preserve fields(0 To 7)
            End If
            valid = (UBound(fields) = 7)
            If valid Then
                For j = 0 To 7
                    If Not IntegerInRange(CStr(fields(j)), -2147483648#, 2147483647#, n) Then
                        valid = False
                        Exit For
                    End If
                    nums(j) = CLng(n)
                Next j
            End If
            If valid Then
                tno = CStr(nums(1))
                valid = (tno = headerTNO) And tnos.Exists(tno)
                valid = valid And nums(4) >= 0 And nums(4) <= 255
                valid = valid And nums(5) >= 0 And nums(5) <= 1023
                valid = valid And nums(6) >= 0 And nums(6) <= 15
                valid = valid And nums(7) >= 0 And nums(7) <= 15
            End If
            If valid Then
                key = CStr(nums(2)) & "," & CStr(nums(3)) & "," & CStr(nums(4)) & "," & CStr(nums(5)) & "," & CStr(nums(6)) & "," & CStr(nums(7))
                If Not addresses.Exists(key) Then
                    Set flags = CreateObject("Scripting.Dictionary")
                    addresses.Add key, flags
                End If
                Set flags = addresses.Item(key)
                flags(tno) = 1
                accepted = accepted + 1
            Else
                skipped = skipped + 1
            End If
        End If
    Next token
    If addresses.Count > dst.Rows.Count - 1 Then Err.Raise vbObjectError + 4, , "Too many unique addresses for one Excel sheet."
    ' Stage all results before replacing the previous output.
    ReDim result(1 To addresses.Count + 1, 1 To 6 + tnos.Count)
    fields = Array("X", "Y", "BLK", "WL", "BL", "IO")
    For j = 0 To 5: result(1, j + 1) = fields(j): Next j
    keys = tnos.Keys
    For j = 0 To tnos.Count - 1: result(1, 7 + j) = CLng(keys(j)): Next j
    r = 2
    For Each v In addresses.Keys
        addr = Split(CStr(v), ",")
        For j = 0 To 5: result(r, j + 1) = CLng(addr(j)): Next j
        Set flags = addresses.Item(v)
        For Each fields In flags.Keys
            col = CLng(tnos.Item(CStr(fields)))
            result(r, col) = 1
        Next fields
        r = r + 1
    Next v
    Application.ScreenUpdating = False
    dst.UsedRange.ClearContents
    dst.Range("A1").Resize(UBound(result, 1), UBound(result, 2)).Value2 = result
    With dst.Range("A1").Resize(1, UBound(result, 2))
        .Interior.Color = RGB(23, 50, 77)
        .Font.Color = vbWhite
        .Font.Bold = True
    End With
    dst.Range("A1").Resize(UBound(result, 1), UBound(result, 2)).NumberFormat = "0"
    dst.Range("A1").Resize(UBound(result, 1), UBound(result, 2)).Columns.ColumnWidth = 12
    cfg.Range("D2").Value2 = path
    cfg.Range("D3").Value2 = Now
    cfg.Range("D4").Value2 = addresses.Count
    cfg.Range("D5").Value2 = accepted
    cfg.Range("D6").Value2 = skipped
    Application.ScreenUpdating = oldScreen
    MsgBox "Import complete. Unique addresses: " & addresses.Count & vbCrLf & "Accepted records: " & accepted & vbCrLf & "Skipped invalid records in selected sections: " & skipped, vbInformation
    Exit Sub
Failed:
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    Application.ScreenUpdating = oldScreen
    MsgBox "Import failed: " & Err.Description, vbExclamation
End Sub

Private Function IntegerInRange(ByVal text As String, ByVal lower As Double, ByVal upper As Double, ByRef number As Double) As Boolean
    Dim re As Object
    On Error GoTo Invalid
    text = Trim$(Replace(Replace(Replace(text, vbTab, " "), vbCr, " "), vbLf, " "))
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^[+-]?[0-9]+$"
    If Not re.Test(text) Then Exit Function
    number = CDbl(text)
    IntegerInRange = (number >= lower And number <= upper)
    Exit Function
Invalid:
    IntegerInRange = False
End Function
