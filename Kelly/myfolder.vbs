'Keep running the script even if it tries to read a key/value that doesn't exist
On Error Resume Next

Dim WshShell, fso, bKey, cn, p1, mybox, entryname, WinFolder
Set WshShell = WScript.CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

WinFolder = fso.GetSpecialFolder(0)

Set fso = Nothing

p1 = "HKEY_CLASSES_ROOT\Directory\shell\"

'Get the name that they want for the entry
t = "Add a Folder to the Start Button Right Click"
cn = InputBox("Enter a descriptive name for the entry. It does not have to " & _
	"be the actual name of the folder." & vbCR & vbCR & "Enter the name of " & _
	"an existing item to change it.", t, n)

'Check for empty string or Cancel
If cn <> "" Then
  entryname = cn
  'Create the registry key string
  p1 = p1 & cn & "\"
  wshshell.RegWrite p1 & "", cn
Else
  'if the string was empty or they clicked Cancel
  WScript.Quit
End If

bKey = WshShell.RegRead(p1 & "command\")

n = wshshell.RegRead(bKey & "")

'Give a default choice
n = "C:\Documents and Settings"

t = "Add a Folder to the Start Button Right Click"
cn = InputBox("Enter the drive and folder to open Explorer in. Documents and Settings" & _
	"\<username> is not recommended in a multi-user system.", t, n)
If cn <> "" Then
  cn = Winfolder & "\EXPLORER.EXE /e," & cn
  wshshell.RegWrite p1 & "command\", cn
Else
  Wscript.Quit
End If

MsgBox entryname & " has been added to the Start Button right click menu.",4096,"Finished!"