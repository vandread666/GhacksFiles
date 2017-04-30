'trillian_add.vbs - Adds Trillian to IE's Tools menu and Toolbar
'© 2003 by Kelly and Doug Knox

On Error Resume Next

'Declare variables
Dim WSHShell, p
Dim GUID, InstallLocation

'Set the Windows Script Host Shell and assign values to variables
Set WSHShell = WScript.CreateObject("WScript.Shell")
GUID = "{2ef50289-0ea7-482e-a30b-4947a81e44cf}"

InstallLocation = InputBox("Enter the path to the Trillian EXE file:","File Location","C:\Program Files\Trillian\")

If InstallLocation <> "" Then

  If Right(InstallLocation,1) <> "\" Then
    InstallLocation = InstallLocation & "\"
  End If

  p = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\Extensions\"

  WshShell.RegWrite P & GUID & "\ButtonText","Trillian"
  WshShell.RegWrite P & GUID & "\clsid","{1FBA04EE-3024-11d2-8F1F-0000F87ABD16}"
  WshShell.RegWrite P & GUID & "\Default Visible","YES"
  WshShell.RegWrite P & GUID & "\Exec", InstallLocation & "Trillian"
  WshShell.RegWrite P & GUID & "\HotIcon",InstallLocation & "Trillian.exe,117"
  WshShell.RegWrite P & GUID & "\Icon",InstallLocation & "Trillian.exe,191"
  WshShell.RegWrite P & GUID & "\MenuText","Trillian"

  Set WshShell = Nothing

  MsgBox "Trillian has been added to Internet" & vbCR & " Explorer's Tools Menu and Toolbar",4096,"Finished!"

Else

  MsgBox "Operation cancelled!", 4096, "User Cancelled!"

End If