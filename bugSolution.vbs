Function MyFunction(param1)
  On Error Resume Next
  'Some code here
  If param1 = "" Then
    Err.Raise 1001, "MyFunction", "param1 cannot be empty"
    If Err.Number <> 0 Then
      'Handle the error appropriately. For example, log it, return a default value or exit gracefully.
      WScript.Echo "Error: " & Err.Description
      Err.Clear
    End If
  End If
  On Error GoTo 0
  'More code here
End Function