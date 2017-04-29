msgbox ReadClipboard

function ReadClipboard
 ' Reads Clipboard into a variable
 ' by Christoph Basedau
 with createobject("internetexplorer.application")
  .navigate "about:blank" 
  ReadClipboard = _
   .document.parentwindow.clipboardData.getData ("text")
  .quit
 end with
end function