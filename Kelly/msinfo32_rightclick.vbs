'msinfo32_rightclick.vbs - Add System Information (MSINFO32.EXE) to the right click

Option Explicit
On Error Resume Next

'Declare variables
Dim WSHShell, MyBox, p1, q1, t
Dim jobfunc, strName

'Set the Windows Script Host Shell and assign values to variables
Set WSHShell = WScript.CreateObject("WScript.Shell")

strName = InputBox("Enter the name you would like for the entry.","Add System Information to right click","System Information")

p1 = "HKEY_CLASSES_ROOT\Directory\shell\MSINFO32\"

q1="%CommonProgramFiles%\Microsoft Shared\MSInfo\msinfo32.exe"

'Describe the funtion of the script for a dialog box

jobfunc = "System Information has been added to the" & vbCR
jobfunc = jobfunc & "right click context menu in Explorer."

'This section writes the correct values to the Registry

WSHShell.RegWrite p1, strName, "REG_SZ"
WshShell.RegWrite p1 & "command\", q1, "REG_EXPAND_SZ"

t = "Confirmation"
MyBox = MsgBox (jobfunc, 4096, t)
Set WshShell = Nothing
