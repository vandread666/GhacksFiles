Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p1, p2, n, cn, vbdefaultbutton
Dim itemtype

p1 = "HKLM\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\"

n = ws.RegRead(p1 & "ComputerName")
t = "Change Computer Name"
cn = InputBox("Type new Name and click OK", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "ComputerName", cn
End If

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		ws.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub
