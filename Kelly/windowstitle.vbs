Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p, p1, n, cn, mybox, itemtype, vbdefaultbutton 

p = "HKCU\Software\Microsoft\Internet Explorer\Main\Window Title"
itemtype = "REG_SZ"
n = "Microsoft Internet Explorer" 
Ws.RegWrite p, n, itemtype

p1 = "HKCU\Software\Microsoft\Internet Explorer\Main\"

n = ws.RegRead(p1 & "Window Title")
t = "Change Provided By or Whole Title for IE"
cn = InputBox("Type New Caption and click OK", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "Window Title", cn
End If

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		ws.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub