On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_UserAccount",,48)
For Each objItem in colItems
    Wscript.Echo "AccountType: " & objItem.AccountType
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Disabled: " & objItem.Disabled
    Wscript.Echo "Domain: " & objItem.Domain
    Wscript.Echo "FullName: " & objItem.FullName
    Wscript.Echo "InstallDate: " & objItem.InstallDate
    Wscript.Echo "Lockout: " & objItem.Lockout
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PasswordChangeable: " & objItem.PasswordChangeable
    Wscript.Echo "PasswordExpires: " & objItem.PasswordExpires
    Wscript.Echo "PasswordRequired: " & objItem.PasswordRequired
    Wscript.Echo "SID: " & objItem.SID
    Wscript.Echo "SIDType: " & objItem.SIDType
    Wscript.Echo "Status: " & objItem.Status
Next
