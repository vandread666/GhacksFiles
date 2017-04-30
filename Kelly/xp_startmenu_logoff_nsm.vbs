'xp_startmenu_logoff_nsm.vbs - Disable/Enable Log Off on XP's Start Menu
'© Doug Knox - rev 04/18/2002
'This code may be freely distributed/modified
'Downloaded from www.dougknox.com

Option Explicit
On Error Resume Next

'Declare variables
Dim WSHShell, n, MyBox, p, p1, t, mustboot, errnum, vers
Dim itemtype
Dim enab, disab, jobfunc

Set WSHShell = WScript.CreateObject("WScript.Shell")
p1 = "HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\StartMenuLogoff"

NewStartMenu = 0
ClassicStartMenu = 0
itemtype = "REG_DWORD"
mustboot = "Log off and back on, or restart your pc to" 
mustboot = mustboot & vbCR & "effect the changes"
enab = "ENABLED"
disab = "DISABLED"
jobfunc = "Start Menu Log off is now "

'This section tries to read the registry key value. If not present an 
'error is generated.  Normal error return should be 0 if value is 
'present
t = "Confirmation"
Err.Clear
n = WSHShell.RegRead (p1)
errnum = Err.Number

if errnum <> 0 then
'Create the registry key value for StartMenuLogoff with value 0
	WSHShell.RegWrite p1, 0, itemtype
End If

'If the key is present, or was created, it is toggled
'Confirmations can be disabled by commenting out 
'the two MyBox lines below

If n = 0 Then
	n = 1
WSHShell.RegWrite p1, n, itemtype
Mybox = MsgBox(jobfunc & disab & vbCR & mustboot, 4096, t)

ElseIf n = 1 then
	n = 0
WSHShell.RegWrite p1, n, itemtype

Mybox = MsgBox(jobfunc & enab & vbCR & mustboot, 4096, t)
End If