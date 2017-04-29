'XP_sendtodesktop.vbs - Repairs the Send To Desktop function
'© Doug Knox - 4/7/2002
'Updated 01/14/04
'Downloaded from www.dougknox.com

Set WshShell = WScript.CreateObject("WScript.Shell")

a = "HKEY_CLASSES_ROOT\CLSID\{9E56BE61-C50F-11CF-9A2C-00A0C90A90CE}\"

WshShell.RegWrite a,""
WshShell.RegWrite a & "NeverShowExt",""

b = "HKEY_CLASSES_ROOT\CLSID\{9E56BE61-C50F-11CF-9A2C-00A0C90A90CE}\InProcServer32\"
WshShell.RegWrite b ,"C:\Windows\System32\Sendmail.dll"
WshShell.RegWrite b & "ThreadingModel","Apartment"

c = "HKEY_LOCAL_MACHINE\Software\CLASSES\CLSID\{9E56BE61-C50F-11CF-9A2C-00A0C90A90CE}\DefaultIcon\"
WshShell.RegWrite c,"C:WINDOWS\EXPLORER.exe,3"

d = "HKEY_LOCAL_MACHINE\Software\CLASSES\CLSID\{9E56BE61-C50F-11CF-9A2C-00A0C90A90CE}\shellex\DropHandler\"
WshShell.RegWrite d,"{9E56BE61-C50F-11CF-9A2C-00A0C90A90CE}"

'Thanks to MVP Kelly for this tip
Return = WshShell.Run("REGSVR32 /S SENDMAIL.DLL")
Return = WshShell.Run("REGSVR32 /S OLE32.DLL")
Return = WshShell.Run("REGSVR32 /S /I SHELL32.DLL")

Set WshShell = Nothing

MsgBox "Send To Desktop has been fixed",4096,"Finished!"