On Error Resume Next

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Window_Placement"