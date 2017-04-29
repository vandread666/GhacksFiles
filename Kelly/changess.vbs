Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p1, n, cn, MyBox 
Dim itemtype

p1 = "HKCU\Control Panel\Desktop\"

n = ws.RegRead(p1 & "SCRNSAVE.EXE")
t = "Change ScreenSaver Entry"
cn = InputBox("Type in the complete path and name of the screensaver.", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "SCRNSAVE.EXE", cn
End If

MyBox = MsgBox("The change should be immediate, if not hit refresh.", vbOKOnly,"Done")
