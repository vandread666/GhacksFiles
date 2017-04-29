Set Sh = WScript.CreateObject("WScript.Shell")
On Error Resume Next
sOStype = Sh.RegRead(_
	"HKLM\SYSTEM\CurrentControlSet\Control\ProductOptions\ProductType")
If Err.Number<>0 Then
	Wscript.Echo " This doesn't appear to be an NT-like operating system;" _
		& vbcrlf & "on Win9x use sysedit or msconfig."
Else
	Set FSO = CreateObject("Scripting.FileSystemObject")
	dirConfPath = "%SYSTEMROOT%\System32\"
	fileset=Array("config.nt","autoexec.nt","drivers\etc\hosts", _
		"drivers\etc\lmhosts","drivers\etc\networks", _
		"drivers\etc\protocol","drivers\etc\services")
	On Error Resume Next
	For Each configfile In fileset
		Sh.Run "Notepad " & dirConfPath & configfile
	Next
End If