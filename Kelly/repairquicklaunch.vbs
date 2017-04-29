'Repair Quick Launch
'Kelly Theriot 7/2004
'Kelly's Korner

Set WshShell = WScript.CreateObject("WScript.Shell")

a = "HKEY_CLASSES_ROOT\CLSID\{D82BE2B0-5764-11D0-A96E-00C04FD705A2}\"
WshShell.RegWrite a,"IShellFolderBand"

b = "HKEY_CLASSES_ROOT\CLSID\{D82BE2B0-5764-11D0-A96E-00C04FD705A2}\InProcServer32\"
WshShell.RegWrite b ,"%SystemRoot%\system32\SHELL32.dll", "REG_EXPAND_SZ"
WshShell.RegWrite b & "ThreadingModel","Apartment"


MsgBox "Quick Launch has been repaired",4096,"Finished!"

Set WshShell = Nothing
