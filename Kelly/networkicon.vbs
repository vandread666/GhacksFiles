On Error Resume Next

set ws = WScript.CreateObject("WScript.Shell")

start = "HKLM\SYSTEM\CurrentControlSet"

classGUID = ws.RegRead(start & "\Enum\Root\MS_NDISWANIP\0000\ClassGUID")
netID = ws.RegRead(start & "\Control\Class\" & classGUID & "\0000\NetCfgInstanceID")

ws.RegWrite start & "\Control\Network\" & classGUID & "\" & netID & "\Connection\ShowIcon", 1, "REG_DWORD"

MsgBox "Show icon in Notification Area when connected is set.", 4096,"Finished"