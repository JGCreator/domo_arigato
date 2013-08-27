#include-once
#cs----
$FileMode:
	0 = Read
	1 = Write (append)
	2 = Write (clear & new)
	8 = Create Directory if not exists
#ce----
Func _CreateLog($ProgramName, $FileMode)
	Global $file = FileOpen($ProgramName & "_Results.txt",  $FileMode)
	If $file = -1 Then
		MsgBox(0, "Error", "Unable to open file.")
		Exit
	EndIf
EndFunc

#cs----
Simple line write of created file.
#ce----
Func _logger($text)
	FileWriteLine($file, $text)
EndFunc

