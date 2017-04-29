Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p1, p2, n, g, cn, cg, vbdefaultbutton
Dim itemtype

p1 = "HKEY_CURRENT_USER\Control Panel\International\"

n = ws.RegRead(p1 & "s1159")
g = ws.RegRead(p1 & "s2359")
t = "Change AM and PM Utility"
cn = InputBox("Type a word of six letters or less to replace AM and click OK", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "s1159", cn
End If

cg = InputBox("Type word of six letters or less to replace PM and click OK.", t, g)
If cg <> "" Then
  ws.RegWrite p1 & "s2359", cg
End If

Dim MyBox

MyBox = MsgBox("The changes will take effect when the clock turns to the next minute.", 64,"Done")

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub