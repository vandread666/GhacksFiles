Option Explicit

Dim WSHShell, n, MyBox, p, itemtype, Title, vbdefaultbutton

Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\CabinetState\"
p = p & "Use Search Asst"
itemtype = "REG_SZ"
n = "no"

WSHShell.RegWrite p, n, itemtype
Title = "The Search Assistant is now disabled." & vbCR
Title = Title & "You may need to Log off/Log on" & vbCR
Title = Title & "For the change to take effect."
MyBox = MsgBox(Title,64,"Finished")

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub