On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_NetworkLoginProfile",,48)
For Each objItem in colItems
    Wscript.Echo "AccountExpires: " & objItem.AccountExpires
    Wscript.Echo "AuthorizationFlags: " & objItem.AuthorizationFlags
    Wscript.Echo "BadPasswordCount: " & objItem.BadPasswordCount
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "CodePage: " & objItem.CodePage
    Wscript.Echo "Comment: " & objItem.Comment
    Wscript.Echo "CountryCode: " & objItem.CountryCode
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Flags: " & objItem.Flags
    Wscript.Echo "FullName: " & objItem.FullName
    Wscript.Echo "HomeDirectory: " & objItem.HomeDirectory
    Wscript.Echo "HomeDirectoryDrive: " & objItem.HomeDirectoryDrive
    Wscript.Echo "LastLogoff: " & objItem.LastLogoff
    Wscript.Echo "LastLogon: " & objItem.LastLogon
    Wscript.Echo "LogonHours: " & objItem.LogonHours
    Wscript.Echo "LogonServer: " & objItem.LogonServer
    Wscript.Echo "MaximumStorage: " & objItem.MaximumStorage
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "NumberOfLogons: " & objItem.NumberOfLogons
    Wscript.Echo "Parameters: " & objItem.Parameters
    Wscript.Echo "PasswordAge: " & objItem.PasswordAge
    Wscript.Echo "PasswordExpires: " & objItem.PasswordExpires
    Wscript.Echo "PrimaryGroupId: " & objItem.PrimaryGroupId
    Wscript.Echo "Privileges: " & objItem.Privileges
    Wscript.Echo "Profile: " & objItem.Profile
    Wscript.Echo "ScriptPath: " & objItem.ScriptPath
    Wscript.Echo "SettingID: " & objItem.SettingID
    Wscript.Echo "UnitsPerWeek: " & objItem.UnitsPerWeek
    Wscript.Echo "UserComment: " & objItem.UserComment
    Wscript.Echo "UserId: " & objItem.UserId
    Wscript.Echo "UserType: " & objItem.UserType
    Wscript.Echo "Workstations: " & objItem.Workstations
Next

