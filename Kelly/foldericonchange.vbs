Option Explicit

Set WSHShell = WScript.CreateObject("WScript.Shell")
Dim WSHShell, n, p, p1, p2, t, cn, itemtype, Mybox, Title, vbdefaultbutton

p = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Icons\"
p = p & "3"
itemtype = "REG_SZ"
n = ""
WSHShell.RegWrite p, n, itemtype

p = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Icons\"
p = p & "4"
itemtype = "REG_SZ"
n = ""
WSHShell.RegWrite p, n, itemtype

p1 = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Icons\"

n = WSHShell.RegRead(p1 & "3")
t = "Change the Folder Icon "
cn = InputBox("Type the Exact Path (follow by example) F:\My Icons\nameoficon.ico", t, n)
If cn <> "" Then
  WSHShell.RegWrite p1 & "3", cn
End If

p1 = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Icons\"

n = WSHShell.RegRead(p1 & "4")
t = "Change the Opened Folder Icon"
cn = InputBox("Type the Exact Path (follow by example) F:\My Icons\nameoficon.ico", t, n)
If cn <> "" Then
  WSHShell.RegWrite p1 & "4", cn
End If

Title = "Your Folder Icons have been changed.  Reboot for the changes to take effect." & vbCR
Title = Title & "" & vbCR
Title = Title & "To set back to default, run the script again and leave the entries blank, reboot." & vbCR
MyBox = MsgBox(Title,64,"Finished")

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub
