Dim WshShell, bKey, cn, p1, mybox
Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.RegWrite "HKCR\Directory\shell\Explore CD-Rom\", "Explore CD-Rom", "REG_SZ"
WshShell.RegWrite "HKCR\Directory\shell\Explore CD-Rom\command\", "c:\windows\EXPLORER.EXE /e,", "REG_SZ" 
bKey = WshShell.RegRead("HKCR\Directory\shell\Explore CD-Rom\command\")

p1 = "HKEY_CLASSES_ROOT\Directory\shell\Explore CD-Rom\command\"

n = wshshell.RegRead(p1 & "")
t = "Add Explore CD-Rom to the Start Button - R/C"
cn = InputBox("Add the Drive letter after: /e, (ex.) c:\windows\EXPLORER.EXE /e,g:", t, n)
If cn <> "" Then
  wshshell.RegWrite p1 & "", cn
End If

MsgBox "Explore CD-Rom has been added to the right click. To test, right click the Start Button.",64,"Finished"

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub