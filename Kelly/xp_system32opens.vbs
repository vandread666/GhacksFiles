On Error Resume Next

Set WshShell = WScript.CreateObject("WScript.Shell")

X = WshShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Run\SB Audigy 2 Startup Menu")

If X <> "" Then

WshShell.RegWrite "HKCU\Software\Microsoft\Windows\CurrentVersion\Run\SB Audigy 2 Startup Menu", "/L:ENG"

X = MsgBox("Finished! The System32 folder should not open on boot for this user.", vbOKOnly, "Done")

Else

MsgBox "This script cannot repair your issue. The expected Registry value was not found.", vbOKOnly, "Finished."

End If

Set WshShell = Nothing
