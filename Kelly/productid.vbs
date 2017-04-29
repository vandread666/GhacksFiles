Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p1, p2, n, cn, vbdefaultbutton
Dim itemtype

p1 = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\"

n = ws.RegRead(p1 & "ProductId")
t = "Read Your Product ID"
cn = InputBox("This is your Product ID Number", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "ProductId", cn
End If

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub
