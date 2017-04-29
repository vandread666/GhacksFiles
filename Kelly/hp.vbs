'© Kelly Theriot
'Kelly's Korner http://www.kellys-korner-xp.com/xp_tweaks.htm


Dim WSHShell, n, MyBox, p, itemtype, Title

Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Cached\"
p = p & "{A4DF5659-0801-4A60-9607-1C48695EFDA9} {000214E6-0000-0000-C000-000000000046} 0x401"
itemtype = "REG_DWORD"
n = 1

WSHShell.RegWrite p, n, itemtype

For Each Process in GetObject("winmgmts:"). _
	ExecQuery ("select * from Win32_Process where name='Verclsid.exe'")
   Process.terminate(0)

Next

MsgBox "Finished." & vbcr & vbcr & "© Kelly Theriot", 4096, "Done"