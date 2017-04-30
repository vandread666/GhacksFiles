'xp_unhide_users.vbs - Unhides an XP User account from the Welcome screen
'© Doug Knox - 05/01/02
'Downloaded from www.dougknox.com

Option Explicit
'On Error Resume Next

Set ws = WScript.CreateObject("WScript.Shell")
Dim ws, t, p1, cn, errnum

p1 = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\SpecialAccounts\UserList\"

t = "UnHide XP User Account"
cn = InputBox("Enter the Username to unhide and click OK", t, "")

If cn <> "" Then
  ws.RegDelete p1 & cn
Err.Clear
errnum = Err.Number

If errnum = 0 Then
   MsgBox cn & "will show on the Welcome Screen",4096,"Done!"
Else
   MsgBox "Please make sure you entered a valid user name",4096,"Error!"
End If

End If

