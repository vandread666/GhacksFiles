Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p1, p2, n, cn, MyBox 
Dim itemtype

p1 = "HKLM\SOFTWARE\Microsoft\MessengerService\Policies\"

n = ws.RegRead(p1 & "IMWarning")
t = "Change Warning Message in Messenger"
cn = InputBox("Type in your preferred message:", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "IMWarning", cn
End If

MyBox = MsgBox("You must Log Off/Log On for the changes to take effect.", vbOKOnly,"Done")
