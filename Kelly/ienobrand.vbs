Dim myShell

Set myShell = CreateObject("WScript.Shell")
myShell.Run("rundll32.exe iedkcs32.dll,Clear")

MsgBox "IE Branding has been removed.", 4096,"Finished"