'lowersmfolderadd.vbs - Adds a folder under the Run command on the
'New XP Start Menu
'© 2003 - Kelly Theriot and Doug Knox

Option Explicit
On Error Resume Next

'Declare variables
Dim WSHShell, p1, t
Dim jobfunc, SMText, SMPath

'Set the Windows Script Host Shell and assign values to variables
Set WSHShell = WScript.CreateObject("WScript.Shell")

SMText = Inputbox("Enter the text you want to appear on the Start Menu.","Entry name","Enter Text Here")

If SMText = "" Then
  WScript.Quit
End If

SMPath = Inputbox("Enter the path to the folder.","Enter Folder path","C:\Documents and Settings")

If SMPATH = "" Then
  WScript.Quit
End If

p1 = "HKCR\CLSID\{2559a1f6-21d7-11d4-bdaf-00c04f60b9f0}\"

WSHShell.RegWrite p1, SMText
WSHShell.RegWrite p1 & "Infotip", SMPath
WSHShell.RegWrite p1 & "DefaultIcon\","%systemroot%\system32\shell32.dll,4"
WSHShell.RegWrite p1 & "InProcServer32\","%systemroot%\system32\shdocvw.dll", "REG_EXPAND_SZ"
WSHShell.RegWrite p1 & "InProcServer32\ThreadingModel","Apartment"
WSHShell.RegWrite p1 & "Instance\CLSID","{3f454f0e-42ae-4d7c-8ea3-328250d6e272}"
WSHShell.RegWrite p1 & "Instance\InitPropertyBag\CLSID","{13709620-C279-11CE-A49E-444553540000}"
WSHShell.RegWrite p1 & "Instance\InitPropertyBag\Command","&Open"
WSHShell.RegWrite p1 & "Instance\InitPropertyBag\method","ShellExecute"
WSHShell.RegWrite p1 & "Instance\InitPropertyBag\Param1",SMPath
WSHShell.RegWrite p1 & "shellex\ContextMenuHandlers\{2559a1f6-21d7-11d4-bdaf-00c04f60b9f0}\",""
WSHShell.RegWrite p1 & "shellex\MayChangeDefaultMenu\",""
WSHShell.RegWrite p1 & "ShellFolder\Attributes",0,"REG_DWORD"
WSHShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\StartMenu\StartPanel\ShowOEMLink\NoOEMLinkInstalled",0,"REG_DWORD"

Set WshShell = Nothing

'Describe the funtion of the script for a dialog box

jobfunc = SMText & vbcr & SMPath & vbcr & "Has been added to the Start Menu" & vbCR & vbcr
jobfunc = jobfunc & "Log off/log on to see the changes."

t = "Confirmation"
MsgBox jobfunc, 4096, t