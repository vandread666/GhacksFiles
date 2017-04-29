Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p1, n, cn, MyBox, vbdefaultbutton 
Dim itemtype

p1 = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon\"
n = ws.RegRead(p1 & "")
t = "Change Recycle Bin Icon"
cn = InputBox("Type in the path to the icon in this format: c:\icons\recyclebin.ico", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "", cn
  
End If

p1 = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon\"

n = ws.RegRead(p1 & "Empty")
t = "Change Recycle Bin Empty Icon"
cn = InputBox("Type in the same path, they must match.", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "Empty", cn
End If

p1 = "HKCR\CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon\"

n = ws.RegRead(p1 & "Full")
t = "Change Recycle Bin Full Icon"
cn = InputBox("Type in the path to the icon in this format: c:\icons\recyclebinfull.ico", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "Full", cn
End If

MyBox = MsgBox("You must Log Off/Log On for the changes to take effect.", 64,"Done")

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub

