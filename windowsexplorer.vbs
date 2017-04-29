Dim WshShell, bKey, cn, p1, mybox
Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.RegWrite "HKCR\Directory\shell\Windows Explorer\", "Windows Explorer", "REG_SZ"
WshShell.RegWrite "HKCR\Directory\shell\Windows Explorer\command\", "EXPLORER.EXE /e,", "REG_SZ" 

bKey = WshShell.RegRead("HKCR\Directory\shell\Windows Explorer\")

p1 = "HKEY_CLASSES_ROOT\Directory\shell\Windows Explorer\"

n = wshshell.RegRead(p1 & "")
t = "Add Windows Explorer to the Start Button Right Click"
cn = InputBox("Type in what you want listed for the name. It does not have to be the actual name Windows Explorer.", t, n)
If cn <> "" Then
  wshshell.RegWrite p1 & "", cn
End If

bKey = WshShell.RegRead("HKCR\Directory\shell\Windows Explorer\command\")

p1 = "HKEY_CLASSES_ROOT\Directory\shell\Windows Explorer\command\"

n = wshshell.RegRead(p1 & "")
t = "Add Windows Explorer to the Start Button Right Click"
cn = InputBox("To have Windows Explorer open closed, use what it there, to have Windows Explorer open exapanded, type in: %windir%\EXPLORER.EXE /e,c:   Also modify c: if that is not the where Windows is installed.", t, n)
If cn <> "" Then
  wshshell.RegWrite p1 & "", cn
End If


MsgBox "The folder of your choice has been added to the right click.  If you left it set to default, it will open to My Computer. To test, right click the Start Button.",64,"Finished"

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub