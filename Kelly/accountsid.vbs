On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_AccountSID",,48)
For Each objItem in colItems
    Wscript.Echo "Element: " & objItem.Element
    Wscript.Echo "Setting: " & objItem.Setting
Next

