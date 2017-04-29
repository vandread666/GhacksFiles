Dim WshShell, bKey, cn, p1, mybox
Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.RegWrite "HKCR\Directory\shell\Send E-Mail\", "Send E-Mail", "REG_SZ"
WshShell.RegWrite "HKCR\Directory\shell\Send E-Mail\command\", "c:\program files\outlook express\msimn.exe /mailurl: ", "REG_SZ" 

MsgBox "Send E-Mail has been added to your right click.  To test, right click the Start Button.",64,"Finished"

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub