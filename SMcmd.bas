Attribute VB_Name = "SMcmd"
Function cm(Optional cell As Range) As Double
    If cell Is Nothing Then
        Set cell = Application.Caller
    End If
    Application.Volatile
    For c = -4 To -2
        If IsNumeric(cell.Offset(-2, c).Value) And Not IsEmpty(cell.Offset(-2, c).Value) Then
            If cm = 0 Then cm = 1
            cm = cm * cell.Offset(-2, c)
        End If
    Next
    For r = -2 To 0
        If IsNumeric(cell.Offset(r, -1).Value) And Not IsEmpty(cell.Offset(r, -1).Value) Then
            If cm = 0 Then cm = 1
            cm = cm * cell.Offset(r, -1)
        End If
    Next
End Function
Function sm(Optional cell As Range) As Double
    If cell Is Nothing Then
        Set cell = Application.Caller
    End If
    Application.Volatile
    For c = -4 To -2
        If IsNumeric(cell.Offset(-1, c).Value) And Not IsEmpty(cell.Offset(-1, c).Value) Then
            If sm = 0 Then sm = 1
            sm = sm * cell.Offset(-1, c)
        End If
    Next
    For r = -1 To 0
        If IsNumeric(cell.Offset(r, -1).Value) And Not IsEmpty(cell.Offset(r, -1).Value) Then
            If sm = 0 Then sm = 1
            sm = sm * cell.Offset(r, -1)
        End If
    Next
End Function
Function m(r As Range) As Double
    Dim cell As Range
    Dim subt As Double
    Application.Volatile
    m = 0
    For Each cell In r
        subt = cell.Value
        For c = -3 To -1
            If IsNumeric(cell.Offset(0, c).Value) And Not IsEmpty(cell.Offset(0, c).Value) Then
                subt = subt * cell.Offset(0, c).Value
            End If
        Next
        m = m + subt
    Next
End Function

