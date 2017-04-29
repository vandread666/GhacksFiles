Option Explicit
On Error Resume Next

Dim WSHShell, n, p, itemtype, MyBox
Set WSHShell = WScript.CreateObject("WScript.Shell")

p = "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\NoStartMenuSubFolders"

itemtype = "REG_DWORD"

n = WSHShell.RegRead (p)
errnum = Err.Number

if errnum <> 0 then

	WSHShell.RegWrite p, 1, itemtype
End If

If n = 0 Then
   WshShell.RegWrite p, 1, itemtype
   MyBox = MsgBox("Start Menu Subfolders are now DISABLED", 64, "Show or Hide Start Menu Subfolders")
End If

If n = 1 Then 
   WshShell.Regwrite p, 0, itemtype
   MyBox = MsgBox("Start Menu Subfolders are now ENABLED", 64, "Show or Hide Start Menu Subfolders")
End If

Set WshShell = Nothing