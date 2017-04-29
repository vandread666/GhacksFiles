Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p1, n, cn, MyBox, vbdefaultbutton 
Dim itemtype

p1 = "HKEY_CURRENT_USER\Control Panel\Colors\"

n = ws.RegRead(p1 & "HilightText")
t = "Change RGB Values"
cn = InputBox("Type in the new values in this format:  0  0 255", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "HilightText", cn
End If

MyBox = MsgBox("You must Log Off/Log On for the changes to take effect.", vbOKOnly,"Done")

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub
