'Restore CD-Roms and DVD's to Explorer
'xp_cd_dvd_fix.vbs
'© Doug Knox - rev 04/14/2002
'This code may be freely distributed/modified
'Downloaded from www.dougknox.com 
'based on cdgone.reg

Option Explicit
On Error Resume Next

Dim WshShell, Message

Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.RegDelete "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Class\{4D36E965-E325-11CE-BFC1-08002BE10318}\UpperFilters"
WshShell.RegDelete "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Class\{4D36E965-E325-11CE-BFC1-08002BE10318}\LowerFilters"
WshShell.RegDelete "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Cdr4_2K\"
WshShell.RegDelete "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Cdralw2k\"
WshShell.RegDelete "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Cdudf\"
WshShell.RegDelete "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\UdfReadr\"
Set WshShell = Nothing

Message = "Your CD/DVD-Rom drives should now appear in Windows Explorer." & vbCR
Message = Message & "You may need to reboot your computer to see the change."

MsgBox Message, 4096,"Finished!"