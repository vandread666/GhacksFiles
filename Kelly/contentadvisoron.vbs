Set WSHShell = CreateObject("WScript.Shell")
RegKey = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\policies\Ratings\.Default\"
WSHShell.RegWrite regkey & "Enabled", 1, "REG_DWORD"

