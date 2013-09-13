#Include <GuiMenu.au3>
#include <GUIConstantsEx.au3>

#include <Array.au3>
;~ #include 'RegEx_Get_CopyBook.au3'


;~ Opt('MustDeclareVars', 1)

_Main()

Func _Main()
	Local $filemenu, $FileOpen, $recentfilesmenu, $separator1, $txtbx
	Local $hGUI, $hFile, $hEdit, $hHelp, $hMain
	Local $exititem, $helpmenu, $aboutitem, $okbutton, $cancelbutton
	Local $msg, $file
	#forceref $separator1

	$Window = GUICreate("GUI menu", 300, 120)
	GUICtrlCreateLabel('Enter or select from File/Open...' & @lf & 'the location of the file you wish to search.',10,5)
	
	Dim $arRecent[1]
	Dim $arValues[1]
	Dim $times = 0
	
	; input box
	$txtbx = GUICtrlCreateInput('',10,45,280,20)
	; buttons
	$okbutton = GUICtrlCreateButton("OK", 50, 75, 70, 20)
	$cancelbutton = GUICtrlCreateButton("Cancel", 180, 75, 70, 20)
	Local Enum $idNew = 1000, $idOpen, $idSave, $idExit, $idCut, $idCopy, $idPaste, $idAbout

	; file menu
	$filemenu = GUICtrlCreateMenu("File")
	
	
	
;~ 	$hFile = _GUICtrlMenu_CreateMenu ()
;~     $FileOpen = _GUICtrlMenu_InsertMenuItem ($hFile, 1, "&Open", $idOpen)
;~ 	ConsoleWrite($FileOpen)
;~     _GUICtrlMenu_InsertMenuItem ($hFile, 3, "", 0)
;~     $exititem = _GUICtrlMenu_InsertMenuItem ($hFile, 4, "E&xit", $idExit)
;~ 	
;~ 	$hHelp = _GUICtrlMenu_CreateMenu()
;~ 		$About = _GUICtrlMenu_InsertMenuItem($hHelp, 0, '&About', $idAbout)
;~ 	
;~ 	$hMain = _GUICtrlMenu_CreateMenu ()
;~ 	_GUICtrlMenu_InsertMenuItem ($hMain, 0, "&File", 0, $hFile)
;~ 	_GUICtrlMenu_InsertMenuItem ($hMain, 2, "&Help", 0, $hHelp)
;~ 	_GUICtrlMenu_SetMenu ($Window, $hMain)
	
	
;~ 	$helpmenu = _GUICtrlMenu_CreateMenu ()
;~     $aboutitem = _GUICtrlMenu_InsertMenuItem ($helpmenu, 0, "&About", $idAbout)

	


	$FileOpen = GUICtrlCreateMenuItem("Open...", $filemenu)
	$separator1 = GUICtrlCreateMenuItem("", $filemenu)
	$exititem = GUICtrlCreateMenuItem("Exit", $filemenu)
	
	; help menu
	$helpmenu = GUICtrlCreateMenu("?")
	$About = GUICtrlCreateMenuItem("About", $helpmenu)
	
	
GUISetState()
	

	; look for user activity
	While 1
		$msg = GUIGetMsg()
		
		Dim $startpath = ''
		
		; determine user action
		Select 
			; user closed
			Case $msg = $GUI_EVENT_CLOSE Or $msg = $cancelbutton or $msg = $exititem
				ConsoleWrite('closing')
				ExitLoop
			; file > open
			Case $msg = $FileOpen
				; start the dialog from reg value saved
				$startpath = RegRead("HKEY_CURRENT_USER\SOFTWARE\USERDEF\RegExCopyBook", "LastSelect")
				
				; open the dialog
				If $startpath <> '' Then
					$file = FileOpenDialog("Choose file...", $startpath, "All (*.*)")
				Else
					$file = FileOpenDialog("Choose file...", 'C:\', "All (*.*)")
				EndIf
				
				; when successful
				If @error <> 1 Then 
					; create recent files item first time only
					If $times = 0 Then
						$recentfilesmenu = GUICtrlCreateMenu("Recent Files", $filemenu, 1)
					EndIf
					$times += 1
					ReDim $arRecent[$times] 
					ReDim $arValues[$times]
					$arRecent[$times - 1] =	GUICtrlCreateMenuItem($file, $recentfilesmenu,0)
					$arValues[$times - 1] = $file
					ConsoleWrite('added array record ' & $arRecent[$times - 1] &  ' @ index ' & $times -1 & @lf)
					
					; put the text in the box
					ControlSetText ( "", "", $txtbx, $file )
				EndIf

			Case $msg = $okbutton
				Dim $Path
				$Path = ControlGetText('','',$txtbx)
				
				; check menu radio setting for show results (Keep setting in reg?)
;~ 				$Show = RegRead("HKEY_CURRENT_USER\SOFTWARE\USERDEF\RegExCopyBook", "ShowFoundCopybooks")
				$Show = 1
				
				$Return = _GetCopyBooks($Path, $Show)
				If Not IsArray($Return) Then
					ContinueLoop
				EndIf
				
				RegWrite("HKEY_CURRENT_USER\SOFTWARE\USERDEF\RegExCopyBook", "LastSelect", "REG_SZ", $Path)
				ControlSetText ( "", "", $txtbx, '' )
				
			Case $msg = $About
				MsgBox(0, "About", "GUI Menu Test")
				
			Case $msg <> 0 and $msg <> $GUI_EVENT_MOUSEMOVE
				Dim $index = _ArraySearch($arRecent, $msg)
				If $index <> -1 Then
					ControlSetText ( "", "", $txtbx, $arValues[$index] )
				EndIf
				ConsoleWrite($msg)
				

		EndSelect
	WEnd

	GUIDelete()

	Exit
EndFunc   ;==>_Main



Func _GetCopyBooks($sWorkSpaceFile = '', $bDisplayResults = True)
	
;~ 	If $bStandAlone <> False
		; if no params supplied prompt for them
;~ 		If ($sWorkSpaceFile = ' ') or ($sWorkSpaceFile = '') Then
;~ 			$sWorkSpaceFile = InputBox("Path to Branch", "Where is the file?" & @lf & "(include the filename + extinsion)")
;~ 			If ($sWorkSpaceFile = '') Then Exit
;~ 		EndIf
;~ 	EndIf
	
	; open the file 
	ConsoleWrite('opening file' & @LF)
	$SearchFile = FileOpen($sWorkSpaceFile, 0)
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
	If $bDisplayResults = True Then
		_ArrayDisplay($arrResults)
	EndIf
	
	Return $arrResults

EndFunc
