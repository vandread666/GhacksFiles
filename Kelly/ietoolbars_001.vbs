Option Explicit
On Error Resume Next
Dim p1, p2, p3, WshShell, MyBox, vbdefaultbutton

p1 ="HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Toolbar\Explorer\ITBarLayout"
p2 = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Toolbar\ShellBrowser\ITBarLayout"
p3 = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Toolbar\WebBrowser\ITBarLayout"

Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.RegDelete p1
WshShell.RegDelete p2
WshShell.RegDelete p3

Set WshShell = Nothing

MyBox = MsgBox("Explorer Toolbars have been restored.", 64, "Finished")

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub