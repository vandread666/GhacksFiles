Option Explicit

Set WSHShell = WScript.CreateObject("WScript.Shell")
Dim WSHShell, n, MyBox, p, itemtype

p = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\"
p = p & "LegalNoticeCaption"
itemtype = "REG_SZ"
n = ""

p = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\"
p = p & "LegalNoticeText"
itemtype = "REG_SZ"
n = ""

WSHShell.RegWrite p, n, itemtype


Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p1, p2, n1, cn 


p1 = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\"

n1 = ws.RegRead(p1 & "LegalNoticeCaption")
t = "Logon, Logoff Message Dialog Boxes"
cn = InputBox("Type Logon Message and click OK", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "LegalNoticeCaption", cn
End If

n1 = ws.RegRead(p1 & "LegalNoticeText")
t = "Logon, Logoff Message Dialog Boxes"
cn = InputBox("Type Logoff Message and click OK", t, n)
If cn <> "" Then
  
  ws.RegWrite p1 & "LegalNoticeText", cn


End If

