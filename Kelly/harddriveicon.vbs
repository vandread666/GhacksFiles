Option Explicit

Set WSHShell = WScript.CreateObject("WScript.Shell")
Dim WSHShell, n, p, p1, p2, t, cn, itemtype, Mybox, Title

p = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Icons\"
p = p & "8"
itemtype = "REG_SZ"
n = ""
WSHShell.RegWrite p, n, itemtype

p1 = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Icons\"

n = WSHShell.RegRead(p1 & "8")
t = "Change Your Hard Drive Icon "
cn = InputBox("Type the Exact Path (follow by example) F:\My Icons\nameoficon.ico", t, n)
If cn <> "" Then
  WSHShell.RegWrite p1 & "8", cn
End If

Title = "Your Hard Drive Icons have been changed. Reboot for the changes to take effect." & vbCR

Title = Title & "" & vbCR
Title = Title & "To set back to default, run the script again and leave the entry blank, reboot." & vbCR
MyBox = MsgBox(Title,4096,"Finished")