On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Account Where LocalAccount = True")
For Each objItem in colItems
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Domain: " & objItem.Domain
    Wscript.Echo "Install Date: " & objItem.InstallDate
    Wscript.Echo "Local Account: " & objItem.LocalAccount
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "SID: " & objItem.SID
    Wscript.Echo "SID Type: " & objItem.SIDType
    Wscript.Echo "Status: " & objItem.Status
Next

