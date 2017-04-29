'xp_hide_users.vbs
'Hide User on the Windows XP Welcome Screen
'© Doug Knox - 05/01/02
'This code may be freely distributed/modified
'Downloaded from www.dougknox.com

Option Explicit
On Error Resume Next

'Declare variables
Dim WSHShell, n, p, itemtype, MyBox, User, Title, Prompt

'set variables
p = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList\"
itemtype = "REG_DWORD"
n = 0

Prompt = "Enter the username you wish to hide."
Title = "Hide User on Welcome screen"
User = InputBox(Prompt, Title,"")

If User = "" Then
	Title = "Error!"
	Prompt = "You left the user name blank."
	MyBox = MsgBox(Prompt,4096,Title)
Else
	p = p & User
	Set WSHShell = WScript.CreateObject("WScript.Shell")
	WSHShell.RegWrite p, n, itemtype
	Title = "Success"
	Prompt = User & " is now hidden on the Welcome screen."	
	
	MyBox = MsgBox(Prompt, 4096, Title)
End If

Set WshShell = Nothing
