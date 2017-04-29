Option Explicit
On Error Resume Next

Dim WshShell, Message

Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.RegDelete "HKLM\SYSTEM\CurrentControlSet\Control\Class\{4D36E980-E325-11CE-BFC1-08002BE10318}\UpperFilters"
WshShell.RegDelete "HKLM\SYSTEM\CurrentControlSet\Control\Class\{4D36E980-E325-11CE-BFC1-08002BE10318}\Lowerfilters"
WshShell.RegDelete "HKLM\SYSTEM\CurrentControlSet\Control\Class\{4D36E967-E325-11CE-BFC1-08002BE10318}\UpperFilters"
WshShell.RegDelete "HKLM\SYSTEM\CurrentControlSet\Control\Class\{4D36E967-E325-11CE-BFC1-08002BE10318}\LowerFilters"
WshShell.RegWrite "HKLM\System\CurrentControlSet\Services\USB\DisableSelectiveSuspend", 1, "REG_DWORD"


Set WshShell = Nothing

Message = "You must reboot your computer to see the change."

MsgBox Message, 4096,"Finished!"