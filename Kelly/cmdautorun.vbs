Option Explicit


Dim WSHShell, n, MyBox, p, itemtype

Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKCU\Software\Microsoft\Command Processor\"
p = p & "AutoRun"
itemtype = "REG_SZ"
n = ""

WSHShell.RegWrite p, n, itemtype

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p1, p2, n2, cn


p1 = "HKCU\Software\Microsoft\Command Processor\"

n2 = ws.RegRead(p1 & "AutoRun")
t = "CMD AutoRun"
cn = InputBox("Enter exact path to batch file and click OK", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "AutoRun", cn
End If