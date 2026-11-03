Public Function BLK_AddressTNOs() As Object
    Dim result As Object, i As Long
    Set result = CreateObject("Scripting.Dictionary")
    For i = LBound(mSequence) To UBound(mSequence)
        result.Add mSequence(i), result.Count
    Next i
    Set BLK_AddressTNOs = result
End Function

Private Sub ReadItemColumns(ByVal data As Variant, ByVal monitor As Boolean, ByVal seen As Object, ByRef tnos() As String, ByRef names() As String, ByRef volts() As Double, ByRef count As Long)
    Dim i As Long, tno As String, label As String, j As Long
    label = IIf(monitor, "Monitor", "Function")
    For i = LBound(data, 1) To UBound(data, 1)
        For j = 1 To 3
            If IsError(data(i, j)) Then Err.Raise vbObjectError + 302, , label & " cell error at row " & i + 1
        Next j
        If Len(Trim$(CStr(data(i, 1)))) > 0 Then
            tno = Trim$(CStr(data(i, 1)))
            If Len(tno) <> 3 Then Err.Raise vbObjectError + 302, , label & " TNO must have 3 digits at row " & i + 1
            tno = CStr(CLng(UnsignedValue(tno, 999)))
            If CLng(tno) < 100 Then Err.Raise vbObjectError + 302, , label & " TNO must be 100-999."
            If seen.Exists(tno) Then Err.Raise vbObjectError + 303, , "Duplicate TNO across Monitor/Function: " & tno
            If Len(Trim$(CStr(data(i, 2)))) = 0 Then Err.Raise vbObjectError + 401, , label & " Func. missing at row " & i + 1
            If Len(Trim$(CStr(data(i, 3)))) = 0 Or Not IsNumeric(data(i, 3)) Or VarType(data(i, 3)) = vbBoolean Then Err.Raise vbObjectError + 304, , label & " voltage missing/invalid at row " & i + 1
            count = count + 1
            tnos(count) = tno: names(count) = Trim$(CStr(data(i, 2))): volts(count) = CDbl(data(i, 3))
            seen.Add tno, True
            If monitor Then mMonitorTNOs.Add tno, True
        ElseIf Len(Trim$(CStr(data(i, 2)))) > 0 Or Len(Trim$(CStr(data(i, 3)))) > 0 Then
            Err.Raise vbObjectError + 304, , label & " Func./voltage without TNO at row " & i + 1
        End If
    Next i
End Sub

Private Function HasMonitor(ByRef tnos() As String, ByVal monitor As Object) As Boolean
    Dim i As Long
    For i = LBound(tnos) To UBound(tnos)
        If monitor.Exists(tnos(i)) Then HasMonitor = True: Exit Function
    Next i
End Function

Private Function BuildItemDUTGroup(ByVal data As Object, ByRef tnos() As String, ByRef volts() As Double, ByVal dut As String, ByVal threshold As Double, ByVal monitor As Object) As Variant
    Dim output() As Variant, onlyTNOs() As String, onlyVolts() As Double, judged As Variant
    Dim values As Variant, xy As Variant, key As String, i As Long, b As Long, n As Long, count As Long
    n = UBound(tnos): xy = Split(dut, "|")
    ReDim onlyTNOs(1 To n): ReDim onlyVolts(1 To n)
    For i = LBound(tnos) To UBound(tnos)
        If monitor.Exists(tnos(i)) Then
            count = count + 1: onlyTNOs(count) = tnos(i): onlyVolts(count) = volts(i)
        End If
    Next i
    ReDim output(1 To n + IIf(count > 0, 1, 0), 1 To 260)
    For i = LBound(tnos) To UBound(tnos)
        output(i, 1) = tnos(i): output(i, 2) = volts(i): output(i, 3) = CLng(xy(0)): output(i, 4) = CLng(xy(1))
        key = tnos(i) & "|" & dut
        If data.Exists(key) Then
            values = data(key)
            For b = 0 To 255: output(i, b + 5) = values(b): Next b
        Else
            For b = 0 To 255: output(i, b + 5) = "Not Found": Next b
        End If
    Next i
    If count > 0 Then
        ReDim Preserve onlyTNOs(1 To count): ReDim Preserve onlyVolts(1 To count)
        judged = BuildDUTGroup(data, onlyTNOs, onlyVolts, dut, threshold)
        For b = 1 To 260: output(n + 1, b) = judged(count + 1, b): Next b
    End If
    BuildItemDUTGroup = output
End Function
