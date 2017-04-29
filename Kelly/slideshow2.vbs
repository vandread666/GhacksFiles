'Special thanks to Michael McDonald

Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
On Error Resume Next

Dim ws, n, cn, p, itemtype, t
p="HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\ShellImageView\"
p = p & "Timeout"
itemtype = "REG_DWORD"
n = ws.RegRead(p)

If Err.number <> 0 Then
  n = 10
Else
  n = left(n, len(n) - 3)
End If

t = "Change Slideshow Photo Display Time"
cn = InputBox("Enter length of time in seconds and click Ok. For the change to take effect in a particular folder, you must close the folder and then re-open it.", t, n)

If cn <> "" Then
  cn = cn & "000"
  ws.RegWrite p, cn, itemtype
End If
