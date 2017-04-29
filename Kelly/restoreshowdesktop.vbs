Dim myShell

Set myShell = CreateObject("WScript.Shell")
myShell.Run("regsvr32 /n /i:U shell32")

Set myShell = Nothing

MsgBox "Toolbar Menu's have been restored.", 4096,"Finished"
