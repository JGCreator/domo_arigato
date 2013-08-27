#cs----
Author:	JGust
Date: 	08/21/13

Description:
	Gets a list of copybooks from any given cobol text file.
	
	_GetCopyBooks( [Directory[, File name]])

Parameters:
	$WorkSpaceFile:(optional)		The complete path of the workspace branch and file (including the file extension).
		
	$DisplayResults:(optional)	An indicator to show the array of results.
			0 = Display
			1 = No Display
	
Returns:
	Success:	Array of resulting copybook names
	Fail:		-1 	= failed to open file
				 1	= failed to read open file
				 2	= EOF after opening file
				 
Notes:
	This function uses a regular espression search to find all the instances of the word "COPY" followed by
	any number of Spaces >=1 followed by an opening ", the name of the copy file, and the closing ".
	
	This function also assumes that the calling system is asking for a search on the .cbl file located in /SOURCE/ 
	of the workspace tree.
				

#ce----

; C:\Acucorp\workdir\jgust_view\adv1020\source\frmInvoicer.cbl
_GetCopyBooks()
#include <array.au3>
Func _GetCopyBooks($WorkSpaceFile = '', $DisplayResults = 0)
	
	; if no params supplied prompt for them
	If ($WorkSpaceFile = ' ') or ($WorkSpaceFile = '') Then
		$WorkSpaceFile = InputBox("Path to Branch", "Where is the file?" & @lf & "(include the filename + extinsion)")
		If ($WorkSpaceFile = '') Then Exit
	EndIf
	
	; open the file 
	ConsoleWrite('opening file' & @LF)
	$SearchFile = FileOpen($WorkSpaceFile, 0)
	ConsoleWrite('searchfile = ' & $searchfile & @LF)
	If $SearchFile = -1 Then
		ConsoleWrite("An error occurred while opening the file. Check that it exists at the location provided." & @LF)
		ConsoleWrite($SearchFilePath)
		Return -1
	EndIf
	
	; read the file and check for errors
	ConsoleWrite('reading file' & @LF)
	$SearchTxt = FileRead($SearchFile)
	
	If @error = 1 Then
		ConsoleWrite("An error occurred while reading the file. Check that the proper permission is given for reading." & @LF)
		ConsoleWrite($SearchFilePath)
		FileClose($SearchFile)
		Return 1
	ElseIf @error = -1 Then
		ConsoleWrite("An error occurred. EOF reached on initial read. Check for file integrity." & @LF)
		ConsoleWrite($SearchFilePath)
		FileClose($SearchFile)
		Return 2
	EndIf
	
	; clost the file after handle received
	FileClose($SearchFile)
	
	; offset start of regex search
	$nOffset = 1

	; use optioin 3 to return an array of all instances of the match
	$arrResults = StringRegExp( $SearchTxt, ' *?COPY *?"(.*?)"', 3, $nOffset)
	
	; display
	If $DisplayResults = 0 Then
		_ArrayDisplay($arrResults)
	EndIf
	
	Return $arrResults

EndFunc