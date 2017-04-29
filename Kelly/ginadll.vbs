On Error Resume Next

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.RegDelete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\GinaDLL"

Message = "Your Windows Logon Screen is restored" & vbCR & vbCR
Message = Message & "You may need to log off/log on, or" & vbCR
Message = Message & "restart for the change to take effect."

X = MsgBox(Message, vbOKOnly, "Done")
Set WshShell = Nothing