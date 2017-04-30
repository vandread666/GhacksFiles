'xp_taskbargrouping.vbs - Change number of windows that are open before XP will start grouping them
'© Doug Knox - July 7, 2001
'This code may be freely distributed/modified

Option Explicit

'Declare variables
Dim WSHShell, n, MyBox, p, p1, t, cn, Caption, itemtype, errnum

'set variables and Shell
Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarGroupSize"
itemtype = "REG_DWORD"
t = "Change Taskbar Grouping Level"

'This section reads the current value if any and puts it in an Input Box
'where you may change it as you like 


On Error Resume Next

n = WSHShell.RegRead(p)

errnum = Err.Number

If errnum <> 0 then   

n = 0
End If
Caption = "Your current Value is: " & n & vbCR & vbCR
Caption = Caption & "A value of 2 will cause the Taskbar buttons to always group." & VBCR & vbCR
Caption = Caption & "Input your new value."
On Error Goto 0

cn = InputBox(Caption, t, n)
If cn <> "" Then
WSHShell.RegWrite p, cn, itemtype
End If

If cn <>"" Then
MyBox = MsgBox("You must log off/log on for the changes to take effect", vbOKOnly,"Done")
End If