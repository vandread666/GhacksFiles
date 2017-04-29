Set EventSet = GetObject("winmgmts:").ExecQuery("select * from Win32_NTLogEvent where EventCode=4097")

if (EventSet.Count = 0) then WScript.Echo "No Events"

for each LogEvent in EventSet
	WScript.Echo "Event Number: " & LogEvent.RecordNumber
	WScript.Echo "Log File: " & LogEvent.LogFile
	WScript.Echo "Type: " & LogEvent.Type 
	WScript.Echo "Message: " & LogEvent.Message
	WScript.Echo "Time: " & LogEvent.TimeGenerated
	WScript.Echo ""
next