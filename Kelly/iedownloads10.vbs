Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\MaxConnectionsPer1_0Server",10,"REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\MaxConnectionsPerServer",10,"REG_DWORD"

Message = "Internet Explorer Downloads have been increased to 10."

X = MsgBox(Message, vbOKOnly, "Done")
Set WshShell = Nothing

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub
