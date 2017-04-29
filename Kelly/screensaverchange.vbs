Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p1, n, cn, MyBox, vbdefaultbutton 
Dim itemtype

p1 = "HKEY_USERS\.DEFAULT\Control Panel\Desktop\"

n = ws.RegRead(p1 & "SCRNSAVE.EXE")
t = "Change ScreenSaver Entry"
cn = InputBox("Type in the name of the screensaver in this format: screensavername.scr", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "SCRNSAVE.EXE", cn
End If

p1 = "HKEY_USERS\.DEFAULT\Control Panel\Desktop\"

n = ws.RegRead(p1 & "ScreenSaveTimeOut")
t = "Change ScreenSaver Entry"
cn = InputBox("Type in the amount of time 60 = 1min, 600 = 10min", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "ScreenSaveTimeOut", cn
End If

MyBox = MsgBox("You must Log Off/Log On for the changes to take effect.", vbOKOnly,"Done")

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub

