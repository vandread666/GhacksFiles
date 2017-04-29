On Error Resume Next
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.RegWrite "HKCR\lnkfile\NeverShowExt", "", "REG_SZ"
WshShell.RegWrite "HKCR\DocShortcut\NeverShowExt", "", "REG_SZ"
WshShell.RegWrite "HKCR\InternetShortcut\NeverShowExt", "", "REG_SZ"
WshShell.RegWrite "HKCR\piffile\NeverShowExt", "", "REG_SZ"
WshShell.RegWrite "HKCR\SHCmdFile\NeverShowExt", "", "REG_SZ"
WshShell.RegWrite "HKCR\ShellScrap\NeverShowExt", "", "REG_SZ"

Dim MyBox

MyBox = MsgBox("You must Log Off/Log On for the changes to take effect.", vbOKOnly,"Done")

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub