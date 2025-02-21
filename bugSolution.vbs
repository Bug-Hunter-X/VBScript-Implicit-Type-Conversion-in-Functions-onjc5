Function f(x)
  Dim result As Double
  If x < 0 Then
    result = -x
  Else
    result = x
  End If
  f = result
End Function
MsgBox f(-5) 