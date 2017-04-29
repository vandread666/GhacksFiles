Set PageFileSet = GetObject("winmgmts:").InstancesOf ("Win32_PageFile")

for each PageFile in PageFileSet
	WScript.Echo PageFile.Name & Chr(13) 
	WScript.Echo " Initial: " & PageFile.InitialSize & " MB"
	WScript.Echo " Max: " & PageFile.MaximumSize & " MB"	

next
