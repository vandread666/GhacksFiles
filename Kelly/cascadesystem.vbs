Option Explicit
On Error Resume Next

Dim WshShell
Dim Winfolder, ShortcutFolder
Dim FSO, f, oShellLink
Dim RegKey

Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = WScript.CreateObject("WScript.Shell")

RegKey = "HKCR\CLSID\{7b5c104f-4cae-45ee-adb6-2c25c0e5c61a}\"

Winfolder = FSO.GetSpecialFolder(0)

ShortcutFolder = Winfolder & "\SysCpl"

Set f = fso.CreateFolder(ShortcutFolder)

On Error Goto 0

set oShellLink = WshShell.CreateShortcut(ShortcutFolder & "\General.lnk")
oShellLink.TargetPath = "CONTROL"
oShellLink.Arguments = "sysdm.cpl,,0"
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "shell32.dll, 15"
oShellLink.Description = "General tab"
oShellLink.WorkingDirectory = Winfolder
oShellLink.Save

set oShellLink = WshShell.CreateShortcut(ShortcutFolder & "\Computer Name.lnk")
oShellLink.TargetPath = "CONTROL"
oShellLink.Arguments = "sysdm.cpl,,1"
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "shell32.dll, 17"
oShellLink.Description = "Computer Name tab"
oShellLink.WorkingDirectory = Winfolder
oShellLink.Save

set oShellLink = WshShell.CreateShortcut(ShortcutFolder & "\Hardware.lnk")
oShellLink.TargetPath = "CONTROL"
oShellLink.Arguments = "sysdm.cpl,,2"
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "utilman.exe"
oShellLink.Description = "Hardware Tab"
oShellLink.WorkingDirectory = Winfolder
oShellLink.Save

set oShellLink = WshShell.CreateShortcut(ShortcutFolder & "\Advanced.lnk")
oShellLink.TargetPath = "CONTROL"
oShellLink.Arguments = "sysdm.cpl,,3"
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "shell32.dll, 57"
oShellLink.Description = "Advanced tab"
oShellLink.WorkingDirectory = Winfolder
oShellLink.Save

set oShellLink = WshShell.CreateShortcut(ShortcutFolder & "\System Restore.lnk")
oShellLink.TargetPath = "CONTROL"
oShellLink.Arguments = "sysdm.cpl,,4"
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "shell32.dll, 68"
oShellLink.Description = "System Restore tab"
oShellLink.WorkingDirectory = Winfolder
oShellLink.Save

set oShellLink = WshShell.CreateShortcut(ShortcutFolder & "\Automatic Updates.lnk")
oShellLink.TargetPath = "CONTROL"
oShellLink.Arguments = "sysdm.cpl,,5"
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "shell32.dll, 46"
oShellLink.Description = "Automatic Updates tab"
oShellLink.WorkingDirectory = Winfolder
oShellLink.Save

set oShellLink = WshShell.CreateShortcut(ShortcutFolder & "\Remote.lnk")
oShellLink.TargetPath = "CONTROL"
oShellLink.Arguments = "sysdm.cpl,,6"
oShellLink.WindowStyle = 1
oShellLink.IconLocation = "shell32.dll, 18"
oShellLink.Description = "Remote tab"
oShellLink.WorkingDirectory = Winfolder
oShellLink.Save

WshShell.RegWrite RegKey,"System Expanded"
WshShell.RegWrite RegKey & "InfoTip","See information about your computer system, " & _
	"and change settings for hardware, performance, and automatic updates."
WshShell.RegWrite RegKey & "SortOrderIndex",71,"REG_DWORD"
WshShell.RegWrite RegKey & "DefaultIcon\","%SystemRoot%\system32\sysdm.cpl,0","REG_EXPAND_SZ"
WshShell.RegWrite RegKey & "InProcServer32\","%SystemRoot%\system32\shdocvw.dll","REG_EXPAND_SZ"
WshShell.RegWrite RegKey & "InProcServer32\ThreadingModel","Apartment"
WshShell.RegWrite RegKey & "Instance\CLSID","{0AFACED1-E828-11D1-9187-B532F1E9575D}"
WshShell.RegWrite RegKey & "Instance\InitPropertyBag\Target",Winfolder & "\SysCpl"
WshShell.RegWrite RegKey & "ShellFolder\Attributes", &HF0000000,"REG_DWORD"
WshShell.RegWrite RegKey & "ShellFolder\WantsFORPARSING",""

WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ControlPanel\NameSpace\" & _
	"{7b5c104f-4cae-45ee-adb6-2c25c0e5c61a}\","Cascading System menu"

Set oShellLink = Nothing
Set WshShell = Nothing
Set FSO = Nothing

Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\don't load\sysdm.cpl","No"

MsgBox "System Expanded has now been added to the Control Panel", 4096,"Create Cascading System Menu"
