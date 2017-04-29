On Error Resume Next
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.RegDelete "HKCR\lnkfile\NeverShowExt"
WshShell.RegDelete "HKCR\DocShortcut\NeverShowExt"
WshShell.RegDelete "HKCR\InternetShortcut\NeverShowExt"
WshShell.RegDelete "HKCR\piffile\NeverShowExt"
WshShell.RegDelete "HKCR\SHCmdFile\NeverShowExt"
WshShell.RegDelete "HKCR\ShellScrap\NeverShowExt"

Dim MyBox

MyBox = MsgBox("You must Log Off/Log On for the changes to take effect.", vbOKOnly,"Done")

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub