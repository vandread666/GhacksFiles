'Kelly Theriot 11-2004
'Kelly's Korner

Option Explicit

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p1, n, cn, MyBox, vbdefaultbutton 
Dim itemtype

p1 = "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp\"

n = ws.RegRead(p1 & "PortNumber")
t = "Change Port Number"
cn = InputBox("Type in the new port number", t, n)
If cn <> "" Then
  ws.RegWrite p1 & "PortNumber", cn
End If

MyBox = MsgBox("Once done hit refresh (F5) for the changes to take effect.", vbOKOnly,"Done")

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub