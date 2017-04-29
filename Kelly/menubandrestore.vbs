'Kelly Theriot and David Candy 8/2004

Dim myShell

Set myShell = CreateObject("WScript.Shell")
myShell.Run("regsvr32 /i shell32")

Set myShell = Nothing

MsgBox "Toolbar Menu's have been restored.", 4096,"Finished"


