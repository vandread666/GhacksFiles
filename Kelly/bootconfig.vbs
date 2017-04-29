Set BootSet = GetObject("winmgmts:").InstancesOf ("Win32_OperatingSystem")

for each Boot in BootSet
	WScript.Echo Boot.Name
next