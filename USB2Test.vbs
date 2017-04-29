On Error Resume Next
'The presence of the following key seems to indicate USB 2
RegKey = "HKLM\SYSTEM\CurrentControlSet\Enum\USB\ROOT_HUB20"

'Attempt to read the key
CreateObject("WScript.Shell").RegRead(RegKey & "\")

If Err Then
  'An error is always returned when there is no default value, so
  'check for specific error message returned for missing key
  If InStr(LCase(Err.Description), "invalid root") > 0 Then
    MsgBox "This computer does not seem to be USB 2 enabled."
  Else
    MsgBox "This computer appears to be USB 2 enabled."
  End If
End If
