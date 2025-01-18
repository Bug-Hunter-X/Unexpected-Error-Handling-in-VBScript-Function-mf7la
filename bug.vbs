Function MyFunction(param1)
  'Some code here
  If param1 = "" Then
    Err.Raise 1001, "MyFunction", "param1 cannot be empty"
  End If
  'More code here
End Function