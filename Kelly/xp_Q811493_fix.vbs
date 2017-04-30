'xp_Q823980_fix.vbs - Prevents Windows Update from prompting you to install
'Hotfix Q811493, after you have already installed it.
'© Doug Knox and Kelly Theriot - 9/6/2003

Set WshShell = CreateObject("WScript.Shell")

WshShell.RegWrite "HKLM\Software\Microsoft\Windows\CurrentVersion\Uninstall\Q811493\","Needed to prevent Win Update from prompting for this install"
WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\HotFix\Q811493\Installed",1,"REG_DWORD"

Set WshShell = Nothing

Message = "Windows Update should no longer prompt" & vbCR
Message = Message & "for Hotfix Q811493 to be installed." & vbCR & vbCR
Message = Message & "© 2003 - Doug Knox and Kelly Theriot "

MsgBox Message, vbOkOnly, "Finished!"

