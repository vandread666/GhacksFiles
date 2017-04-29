'IE_password_prompt.vbs - Restore the Prompt to save passwords in IE
'© Doug Knox - dbknox@mediaone.net - 02/15/2002
'This code may be freely distributed/modified
'Downloaded from www.dougknox.com
'Inspired by MS-MVP Kelly Theriot

p1 = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FormSuggest PW Ask"

Set WshShell = WScript.CreateObject("WScript.Shell")

WshShell.RegWrite p1,"yes"

Set WshShell = Nothing

x = MsgBox("Internet Explorer will now prompt for passwords.",4096,"Finished")