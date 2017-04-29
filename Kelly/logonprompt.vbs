Option Explicit

Set WSHShell = WScript.CreateObject("WScript.Shell")
Dim WSHShell, n, MyBox, p, itemtype

p = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\"
p = p & "LogonPrompt"
itemtype = "REG_SZ"
n = ""

WSHShell.RegWrite p, n, itemtype

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p1, n2, cn 


p1 = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\"

n2 = ws.RegRead(p1 & "LogonPrompt")
t = "Personalized Message"
cn = InputBox("Type the New Message to appear above the User Name and Password and click OK", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "LogonPrompt", cn
End If