Option Explicit
On Error Resume Next

Dim WSHShell, n, p, itemtype, MyBox
Set WSHShell = WScript.CreateObject("WScript.Shell")

p = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\EnableAutoTray"

itemtype = "REG_DWORD"

n = WSHShell.RegRead (p)
errnum = Err.Number

if errnum <> 1 then

	WSHShell.RegWrite p, 0, itemtype
End If

If n = 0 Then
   WshShell.RegWrite p, 1, itemtype
   MyBox = MsgBox("Show Inactive Icons in the Notification Area are now ENABLED", 64, "Show or Hide Icons in the Notification Area")
End If

If n = 1 Then 
   WshShell.Regwrite p, 0, itemtype
   MyBox = MsgBox("Show Inactive Icons in the Notification Area are now DISABLED", 64, "Show or Hide Icons in the Notification Area")
End If

Set WshShell = Nothing