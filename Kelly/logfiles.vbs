Set LogFileSet = GetObject("winmgmts:").InstancesOf ("Win32_NTEventLogFile")

for each Logfile in LogFileSet
	WScript.Echo " Log Name: " & Logfile.LogfileName & Chr(13), _
		"Number of Records: " & Logfile.NumberOfRecords & Chr(13), _
		"Max Size: " & Logfile.MaxFileSize & " bytes" & Chr(13), _
		"File name: " & Logfile.Name
next
