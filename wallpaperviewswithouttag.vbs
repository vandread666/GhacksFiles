'David Candy and Kelly Theriot 2004

Set objShell = CreateObject("Shell.Application")


objshell.ShellExecute "regsvr32.exe", "/i /s shell32"
objshell.ShellExecute "regsvr32.exe", "/i /s MSHTML.DLL"
objshell.ShellExecute "regsvr32.exe", "/i /s SHIMGVW.DLL"
