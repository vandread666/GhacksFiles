'Always Hide Extension
'alwayshideext.vbs
'Copyright 2003 - Kelly Theriot and Doug Knox

On Error Resume Next

Dim WshShell, strValue, p1, Example

Example = "BMP"


Set WSHShell = WScript.CreateObject("WScript.Shell")

strValue = InputBox("Enter the file extension to hide. Example: BMP","Enter File Extension", Example)
  
If strValue = "" Then
  WScript.Quit
End If

Example = strValue

p1 = WshShell.RegRead("HKCR\." & strValue & "\")

If Err.Number <> 0 Then

  MsgBox "Check the Extension you entered.  It may not be valid.",4096,"Error"

Else

  WshShell.RegWrite "HKCR\" & p1 & "\NeverShowExt", "", "REG_SZ"
  MsgBox strValue & " extensions will always be hidden.  Log off/log on" & vbCR & "or reboot for the changes to take effect.",4096,"Finished"

End If

Set WshShell = Nothing

