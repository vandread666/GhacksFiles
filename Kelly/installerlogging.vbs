Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\Installer\Logging","voicewarmup","REG_SZ"

Message = "Windows Installer Logging is now enabled."

X = MsgBox(Message, vbOKOnly, "Done")
Set WshShell = Nothing