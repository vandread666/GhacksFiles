Dim WshShell, bKey, cn, p1, mybox
Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.RegWrite "HKEY_CLASSES_ROOT\CLSID\{801eb159-9ede-4dcb-919f-b55376cd4ceb}\", "", "REG_SZ"
WshShell.RegWrite "HKEY_CLASSES_ROOT\CLSID\{801eb159-9ede-4dcb-919f-b55376cd4ceb}\DefaultIcon\", "", "REG_SZ" 
WshShell.RegWrite "HKEY_CLASSES_ROOT\CLSID\{801eb159-9ede-4dcb-919f-b55376cd4ceb}\DefaultIcon\InProcServer32\", "shell32.dll", "REG_SZ" 
WshShell.RegWrite "HKEY_CLASSES_ROOT\CLSID\{801eb159-9ede-4dcb-919f-b55376cd4ceb}\DefaultIcon\InProcServer32\ThreadingModel", "Apartment", "REG_SZ" 
WshShell.RegWrite "HKEY_CLASSES_ROOT\CLSID\{801eb159-9ede-4dcb-919f-b55376cd4ceb}\Shell\Open My Folder\Command\", "explorer /root,", "REG_SZ"
WshShell.RegWrite "HKEY_CLASSES_ROOT\CLSID\{801eb159-9ede-4dcb-919f-b55376cd4ceb}\ShellEx\PropertySheetHandlers\","{e33898de-6302-4756-8f0c-5f6c5218e02e}", "REG_SZ"
WshShell.RegWrite "HKEY_CLASSES_ROOT\CLSID\{801eb159-9ede-4dcb-919f-b55376cd4ceb}\ShellFolder\Attributes", "60", "REG_DWORD"
WshShell.RegWrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{801eb159-9ede-4dcb-919f-b55376cd4ceb}\", "", "REG_SZ"
WshShell.Regwrite "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{801eb159-9ede-4dcb-919f-b55376cd4ceb}\", "", "REG_SZ"

bKey = WshShell.RegRead("HKEY_CLASSES_ROOT\CLSID\{801eb159-9ede-4dcb-919f-b55376cd4ceb}\")
p1 = "HKCR\CLSID\{801eb159-9ede-4dcb-919f-b55376cd4ceb}\"

n = wshshell.RegRead(p1 & "")
t = "Add a Folder to My Computer and Desktop"
cn = InputBox("Type Folder Name", t, n)
If cn <> "" Then
  wshshell.RegWrite p1 & "", cn
End If

bKey = WshShell.RegRead("HKEY_CLASSES_ROOT\CLSID\{801eb159-9ede-4dcb-919f-b55376cd4ceb}\DefaultIcon\")
p1 = "HKEY_CLASSES_ROOT\CLSID\{801eb159-9ede-4dcb-919f-b55376cd4ceb}\DefaultIcon\"

n = wshshell.RegRead(p1 & "")
t = "Add a Folder to My Computer and Desktop"
cn = InputBox("Type full path of icon to be used, followed by a ,0 (comma then a zero) with no spaces.", t, n)
If cn <> "" Then
  wshshell.RegWrite p1 & "", cn
End If

bKey = WshShell.RegRead("HKEY_CLASSES_ROOT\CLSID\{801eb159-9ede-4dcb-919f-b55376cd4ceb}\Shell\Open My Folder\Command\")
p1 = "HKEY_CLASSES_ROOT\CLSID\{801eb159-9ede-4dcb-919f-b55376cd4ceb}\Shell\Open My Folder\Command\"

n = wshshell.RegRead(p1 & "")
t = "Add a Folder to My Computer and Desktop"
cn = InputBox("Type full path of Folder after: explorer /root, (example: explorer /root,f:My Stuff).", t, n)
If cn <> "" Then
  wshshell.RegWrite p1 & "", cn
End If

bKey = WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{801eb159-9ede-4dcb-919f-b55376cd4ceb}\")
p1 = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\MyComputer\NameSpace\{801eb159-9ede-4dcb-919f-b55376cd4ceb}\"

n = wshshell.RegRead(p1 & "")
t = "Add a Folder to My Computer and Desktop"
cn = InputBox("Type in Folder Name again.  This is needed for NameSpace.", t, n)
If cn <> "" Then
  wshshell.RegWrite p1 & "", cn
End If

bKey = WshShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{801eb159-9ede-4dcb-919f-b55376cd4ceb}\")
p1 = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{801eb159-9ede-4dcb-919f-b55376cd4ceb}\"

n = wshshell.RegRead(p1 & "")
t = "Add a Folder to My Computer and Desktop"
cn = InputBox("Type in Folder Name once more.  This is needed for NameSpace.", t, n)
If cn <> "" Then
  wshshell.RegWrite p1 & "", cn
End If

MsgBox "The folder of your choice has been added to My Computer and the Desktop.  Once done, open My Computer, it will be listed as a System Folder under Other.",64,"Finished"

VisitKelly's Korner

Sub VisitKelly's Korner
	If MsgBox("This script came from the Tweaks Section of Kelly's Korner" & vbCRLF & vbCRLF & "Would you like to visit Kelly's Web Site now?", vbQuestion + vbYesNo + vbDefaultButton, "Visit Kelly's Korner") =6 Then
		wshshell.Run "http://www.kellys-korner-xp.com/xp_tweaks.htm"
	End If
End Sub