Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p, n, cn, MyBox 
Dim itemtype

p = "HKCU\Software\Policies\Microsoft\Internet Explorer\"
p = p & "Windows Update Menu Text"
itemtype = "REG_SZ"
n = ""
ws.RegWrite p, n, itemtype

p = "HKCU\Software\Policies\Microsoft\Internet Explorer\"

n = ws.RegRead(p & "Windows Update Menu Text")
t = "Change Text under Tools for Windows Update"
cn = InputBox("Type in the text you want to appear under Tools in Internet Explorer. NOTE:  Leaving it empty removes the entry from Tools. To set back to default type in: Windows Update.", t, n)
If cn <> "" Then
  ws.RegWrite p & "Windows Update Menu Text", cn
End If

MyBox = MsgBox("You must re-open Internet Explorer for the changes to take effect.", 64,"Your Changes Have Been Made")
