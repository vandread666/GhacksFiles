Option Explicit

Dim WSHShell, n, MyBox, p, itemtype, Title

Set WSHShell = WScript.CreateObject("WScript.Shell")
p = "HKEY_CLASSES_ROOT\SystemFileAssociations\image\ShellEx\ContextMenuHandlers\ShellImagePreview\"
p = p & ""
itemtype = "REG_SZ"
n = ""

WSHShell.RegWrite p, n, itemtype
Title = "Picture and Fax Viewer is now Disabled." & vbCR
Title = Title & "You may need to Log Off/Log On" & vbCR
Title = Title & "For the change to take effect."
MyBox = MsgBox(Title,4096,"Finished")


