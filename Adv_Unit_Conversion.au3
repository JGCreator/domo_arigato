; Comment ===============================================================================
; Description:  Automate the units creation process for new dealership conversions.
; 
; Requirements: A pre-defined format
; =======================================================================================



; run as admin
#RequireAdmin

#include './Includes/Unit_Conversion_constants.au3'
#include './Includes/CommonControlFunctions.au3'
#include <ListboxConstants.au3>
#include <EditConstants.au3>
#include <WindowsConstants.au3>
#include <GUIConstantsEx.au3>
#Include <GuiListBox.au3>
#include <SendMessage.au3>
#include <Array.au3>

opt('WinTitleMatchMode', 1)
opt('WinTextMatchMode', 2)
opt('GUIOnEventMode',1)
opt('TrayIconDebug', 1)

;~ xml = ObjCreate('System.Xml.XmlTextReader')

;~ obj


_Main()


Func _Main()
	; general global vars
	Global $file_is_open = False	; flag to try opening file on go button if file/open menu not used
	Global $h_file	; handle of file when it is open
	Global $h_AcuWin	; handle of the Units Maintenance window
;~ 	Global $L_count		; a counter for the list box
	
	; debug use: create this key in registry and populate with path to file 
	; to bypass the steps of selecting files from the dialog and use key path by default.
	Global $k = RegRead('HKEY_CURRENT_USER\SOFTWARE\USERDEF\temp', 'temp-key')
	
	
	; Create the window
	Global $Window = GUICreate("Units Conversion",265,270,0,0,$GUI_SS_DEFAULT_GUI,$WS_EX_ACCEPTFILES)

	
	; Create the label and entry field
	GUICtrlCreateLabel("Conversion Data File Path:", 5,5)
	Global $h_file_path = GUICtrlCreateInput('',5,20,200,20)
	GUICtrlSetState($h_file_path,$GUI_DROPACCEPTED)
	
	if $k <> '' Then	; use debug key path if present.
		ControlSetText($Window,'',$h_file_path,$k)
	EndIf
	
	; create the file menu, open.. menu item, and Clear List menu item.
	Local $filemenu, $FileOpen
	$filemenu = GUICtrlCreateMenu("File")
	$FileOpen = GUICtrlCreateMenuItem("Open...", $filemenu)
	$FileClear = GUICtrlCreateMenuItem('Refresh', $filemenu)
	$helpMenu = GUICtrlCreateMenu('?')
	$helpItem = GUICtrlCreateMenuItem('What the..?',$helpMenu)
	
	; create the go button
	$h_GO = GUICtrlCreateButton("Go -->", 210,20,50,20)
	
	; create the list box for results. 
	; List box results are not always the same as console results.
;~ 	Global $h_list = GUICtrlCreateList("Log:", 5,45,255,210,0x4080)
	Global $h_list = GUICtrlCreateList("Log:", 5,45,255,210,0x00204080)
	
	; create dummy keys for hidden delete function
;~ 	Global $h_Del = GUICtrlCreateDummy()
;~ 	dim $aAccel[1][2] 
;~ 	$aAccel[0][0] = '^!{DELETE}' 
;~ 	$aAccel[0][1] = $h_Del
;~ 	GUISetAccelerators($aAccel)
	
	; assign events and functions to perform to the items on the window.
	GUICtrlSetOnEvent($h_GO, "_GoClicked")
	GUICtrlSetOnEvent($FileOpen, "_openDialog")
	GUICtrlSetOnEvent($FileClear, "_clearList")
	GUICtrlSetOnEvent($helpItem,'_whatThe')
	HotKeySet('^!{d}', 'DELETE')
	GUISetOnEvent($GUI_EVENT_CLOSE, "_exit")
	
	; show the window with focus in the path field.
	GUISetState(@SW_SHOW, $window)
	ControlFocus('','',$h_file_path)
	
	while 1 > 0
		sleep( 1000 )
	WEnd
	
EndFunc


Func _exit()
	Exit
EndFunc

Func _openDialog()
	
	$file_path = FileOpenDialog("Choose file...", 'C:\', "All (*.csv)", 10)
	If $file_path <> '' Then
		ControlSetText('','', $h_file_path, $file_path)
	EndIf
	$h_file = FileOpen($file_path,1)
	If $h_file <> -1 Then
		_wCL('file is open')
		$file_is_open = True
	EndIf
	
EndFunc

Func _clearList()
	GUICtrlSetData($h_list, '')
	ConsoleWrite('data cleared.' & @lf & @lf)
	GUICtrlSetData($h_list,'Log:')
	If $file_is_open Then
		FileClose($h_file)
	EndIf
	$h_file = FileOpen(controlgettext($Window,'',$h_file_path),1)
	If $h_file <> -1 Then 
		$file_is_open = True 
		_wCL('file is open')
	EndIf
EndFunc

Func _GoClicked()
	$file_path = ControlGetText('','',$h_file_path)
	If $file_path = '' Then
		return 1
	EndIf
	
	If $file_is_open = False Then
		$h_file = FileOpen($file_path,1)
		If $h_file <> -1 Then
			$file_is_open = True
			_wCL('file is open')
		Else
			ConsoleWrite('File corruption on data file. Can not open.' & @lf)
			MsgBox(0,"Corrupt File!", 'File corruption on data file. Can not open.' & @lf)
			return 1
		EndIf
	EndIf
	
	
	_wCL('waiting for window and focus')
	$result = _getFocus()
	while $result <> 0
		return 2 	; user not ready
	WEnd
	
	Global $splash = SplashTextOn('Unit Conversion Script', _ 
		'The automated unit conversion script is running.' & @LF & _ 
		'If you did not initiate this program, Do Not Interrupt This Process.',400,150)
					
	_wCL('Creating Units:')
	_wCL('  successfully created:')
	
	; read line 1 for headers
	$line = FileReadLine($h_file,1)
	If @error = -1 Then
		_wCL('EOF reached, line 1')
		return 3
	EndIf
	If @error = 1 Then
		_wCL('Error Reading File')
		return 3
	EndIf
	
	; split line into array of headers
	Global $arHeader = StringSplit($line,',',2)
	
	If UBound($arHeader) < 9 Then
		return 3
	EndIf
	
	; csv file doesn't have headers (because someone deleted them)
	; so call the create function on this line
	$result = ''
	If $arHeader[0] <> 'Unit#' Then
		$result = _createUnit($arHeader)
	EndIf
	
	
	; loop through the rest of the file
	
	While 1 > 0 
		BlockInput(1)
		$result = _createUnit()
		
		Select 
			Case $result = 0
				ConsoleWrite('create result: ' & $result & ' = success' & @lf & @lf)
			case $result = -1 
				_wCL('Corrupt file, exiting process.')
				ExitLoop
			Case $result = -2
				ExitLoop
			Case $result = -3
				ConsoleWrite('skipping blank line' & @lf)
				ContinueLoop
			Case $result = 101
				_wCL('EOF reached')
				SplashOff()
				return 0
		EndSelect
	WEnd
	FileClose($h_file)
	SplashOff()

EndFunc

Func _getFocus()
	
	If WinExists('Unit Maintenance') Then
		WinActivate('Unit Maintenance')
		$h_AcuWin = WinGetHandle('Unit Maintenance')
	Else
		While Not WinExists('Unit Maintenance')
	
			MsgBox(0,'Open Unit Maintenance','Please Open the Units Maintenance Program,' & @lf & _ 
				'and start a new session. When finished, put focus on the Unit # field.')
		WEnd
		$h_AcuWin = WinGetHandle('Unit Maintenance')	
	EndIf
	
	If $h_AcuWin = '' Then
		$h_AcuWin = WinGetHandle('Unit Maintenance')
	EndIf
	
	$focus = False
	ConsoleWrite('focus False' & @lf)		
	While $focus = False 
		$Control = ControlGetFocus ($h_AcuWin)
		
		If $Control = $h_unit_nbr Then
			$ready = MsgBox(4,'Ready to Continue?', '')
			If $ready = 7 Then
				ConsoleWrite('user not ready, starting again.' & @lf)
				return 1
			Else
				ConsoleWrite('focus True - $Control: ' & $Control & @lf)
				$focus = True
				return 0
			EndIf
		EndIf
		
		If Not WinExists($h_AcuWin) Then 
			_wCL('Lost Window Handle')
			Return -1
		EndIf
		
	WEnd

EndFunc

Func _createUnit($array = '')
	; check for window still exists
	If Not WinExists($h_AcuWin) Then 
		BlockInput(0)
		_wCL('Lost Window Handle')
		Return -2
	EndIf
	
	; read data from csv and populat correct fields for screen entry
	If Not IsArray($array) Then 
		$line = FileReadLine( $h_file )
		If @error = -1 Then
			BlockInput(0)
			return 101
		EndIf
		$arValue = StringSplit($line, ',',2)
		for $i =0 to UBound($arValue)-1
			ConsoleWrite('sub ' & $i & '= ' & $arValue[$i] & @lf)
		Next
		If ubound($arValue) < 9 Then
			BlockInput(0)
			return -1
		EndIf	
	Else
		$arValue = $array
	EndIf
	
	; skip blank lines
	If $arValue[0] = '' Or $arValue[0] = ' ' Then
		BlockInput(0)
		return -3
	EndIf
	
	ControlSetText('','',$splash, 'Processing: ' & $arValue[0])
	 
	
	; populate the unit number
	_EnterTheText($h_AcuWin,'', $h_unit_nbr, $arValue[0], 4)
;~ 	BlockInput(1) ; block user input
		sleep(250)	
		_SendTab( '{TAB}', 2 )
;~ 	BlockInput(0) ; unlock keyboard and mouse
	
	; wait to see if there is an error
	$count = 0
	While Not WinExists('Unit Maintenance', 'Unit is not on file.')
		sleep(250)
		$count += 1
		If $count = 8 Then		; wait for 3 seconds
			$check = ControlGetText($h_AcuWin,'',$h_unit_type)
			If $check <> '' Then
				_wCL($arValue[0] & 'failed: Unit may exist.')
				BlockInput(0)
				return -2
			EndIf
			BlockInput(0)
			$ans = MsgBox(4,'Keep Waiting?','')
			If $ans = 6 Then
				BlockInput(1)
				$count = 0
			Else
				return -1
			EndIf
		EndIf
	WEnd
	
	ConsoleWrite('Clicking the "Yes" button' & @lf)
	_ClickTheButton('Unit Maintenance', 'Unit is not on file', $h_yes, .5)
	
	; populate the unit type
;~ 	BlockInput(1) ; block the user
		_EnterTheText($h_AcuWin,'', $h_unit_type, $arValue[1],2)
		_SendTab( '{TAB}', 2 )
		sleep( 250 )
		$stat = WinGetText($h_AcuWin)
		$stat = StringInStr($stat, 'INVALID TYPE ENTERED')
;~ 	BlockInput(0) ; unlock keyboard and mouse
	If $stat <> 0 Then
		ConsoleWrite('@error: ' & @error & @lf)
		_wCL($arValue[0] & ' failed: Unit Type error catching found')
		_wCL(chrw(9) & 'Unit Type: ' & $arValue[1])
		BlockInput(0)
		return -2
	EndIf
	
	; populate the unit description
	_EnterTheText($h_AcuWin,'', $h_description, $arValue[2],2)
	; populate the unit cost
	_EnterTheText($h_AcuWin,'', $h_cost, $arValue[4],2)
	
	; populate the unit vin, and check for validation (error)
	sleep(100)
	do 
		ControlFocus($h_AcuWin,'', $h_vin)
		$focus = ControlGetFocus($h_AcuWin)
	Until $focus = $h_vin
;~ 	BlockInput(1) ; block user input
		_SendTab($arValue[3],.25)
		_SendTab( '{TAB}', 2 )
		sleep( 250 )
		$stat = WinGetText($h_AcuWin)
		$stat = StringInStr($stat, 'Invalid VIN - Verify VIN #')
;~ 	BlockInput(0) ; unlock keyboard and mouse
	If $stat <> 0 Then
		ConsoleWrite('@error: ' & @error & @lf)
		_wCL($arValue[0] & ' failed: VIN error catching found:')
		_wCL(chrw(9) & 'VIN: ' & $arValue[3])
		BlockInput(0)
		return -2
	EndIf
	
	; if there is value in list, enter it
	If $arValue[6] <> '' Or $arValue[6] <> 0 Then
		ConsoleWrite('putting list ' & $arValue[6] &@lf)
		_EnterTheText($h_AcuWin,'', $h_list, $arValue[6],.25)
	EndIf
	
	; if there is value in holdback, enter it
	If $arValue[7] <> '' Or $arValue[7] <> 0 Then
		ConsoleWrite('putting holdback' &@lf)
		_EnterTheText($h_AcuWin,'', $h_holdback, $arValue[7],.25)
	EndIf
	
	; if there is value in pack, enter it
	If $arValue[8] <> '' Or $arValue[8] <> 0 Then
		ConsoleWrite('putting pack' &@lf)
		_EnterTheText($h_AcuWin,'', $h_dealer_pack, $arValue[8],.25)
	EndIf
	
	; click and move to the accounting tab
	ConsoleWrite('clicking "Accounting" tab' & @lf)
	ControlClick($h_AcuWin,'',$h_tabs, 'LEFT', 1, $h_acct_x, $h_acct_y)
	
	; put focus on the transaction description and tab to grid
	ConsoleWrite('waiting for focus on transaction description' & @lf)
	Do
		ControlFocus($h_AcuWin,'',$h_trans_desc)
		$focus = ControlGetFocus($h_AcuWin)
		sleep(100)
	Until $focus = $h_trans_desc
	
	If ControlFocus($h_AcuWin,'', $h_trans_desc) Then
;~ 		BlockInput(1) ; prevent user interaction 
			_SendTab('{TAB 2}',.15)
			_SendTab('{DOWN}',.15)
			
			$count = 0
			ConsoleWrite('clearing accounting grid (rows >= 2)')
			Do
				_SendTab('^c',.05)
				$clip = ClipGet()
				_SendTab('{SPACE}',.15)
				_SendTab('{TAB}',.15)
				_SendTab('+{TAB}',.15)
			Until $clip = ' '
	
		_SendTab($arValue[5],.25)
		_SendTab('{TAB 3}',.25)
		$neg_cost = $arValue[4] - ($arValue[4]*2)	
		_SendTab($neg_cost,.25)
		_SendTab('{TAB}',.25)
		_SendTab($arValue[0],.25)
		_SendTab('{TAB}',.25)
;~ 		_SendTab('DNU',.25)
;~ 		BlockInput(0) ; unlock keyboard and mouse
	EndIf
	
	ControlClick($h_AcuWin,'',$h_save_new)
	
	ConsoleWrite('waiting for unit number focus' & @lf)
	Do
		$focus = ControlgetFocus($h_AcuWin)
		sleep( 250 )
	Until $focus = $h_unit_nbr
	
	_wCL(chrw(9) & ': ' & $arValue[0])
	BlockInput(0)
		
EndFunc

Func _wCL($data)	; write the console and the list
	
	GUICtrlSetData($h_list, $data)
	ConsoleWrite($data & @LF)
;~ 	$L_count = _SendMessage($h_list,0x018B,0,0)
	$count = _GUICtrlListBox_GetCount($h_list)

	ConsoleWrite('count: ' & $count & @lf)
;~ 	_SendMessage($h_list, 0x197, $L_count - 1,0)
	If $count > 15 Then
		$curTop = _GUICtrlListBox_GetTopIndex($h_list)
		_GUICtrlListBox_SetTopIndex($h_list,$curTop +1)
	EndIf
	
EndFunc

Func _whatThe()
	ShellExecute("http://goo.gl/WQ8090")
EndFunc

Func DELETE()
	$ans = MsgBox(4,'Are You Sure?!', '')
	
	If $ans = 7 Then
		ConsoleWrite('answered no' &@lf)
		Return -1
	EndIf
	
	$pass = InputBox('Simon Says...', ' ','','#',300,100)
	If $pass <> 'greenlight' Then
		return -1
	EndIf
	
	$ready = _getFocus()
	If $ready = 1 or $ready = -1 Then
		return -1
	EndIf
	BlockInput(1)
	
	$file = FileOpen(controlgettext($window,'',$h_file_path),1)
	
	If Not WinExists($h_AcuWin) Then
		ConsoleWrite('win not existing' &@lf)
		BlockInput(0)
		Return -1
	Else
		WinActivate($h_AcuWin)
	EndIf
	
	
	
	
	ConsoleWrite('reading file'&@lf)
	$ar = FileReadline($file,2)	
	Do
		ConsoleWrite('file read: '&$ar&@lf)
		do 
			ConsoleWrite('getting focus'&@lf)
			ControlFocus($h_AcuWin,'',$h_unit_nbr)
			$focus = ControlGetFocus($h_AcuWin)
		Until $focus = $h_unit_nbr
		ConsoleWrite('have focus'&@lf)
		ConsoleWrite('splitting string'&@lf)
		$ar = StringSplit($ar,',',2)
		ConsoleWrite('entering unit number'&@lf)
		_EnterTheText ( $h_AcuWin, '', $h_unit_nbr, $ar[0], 1)
		_SendTab('{TAB}',1)
		Sleep(500)
		
		ConsoleWrite('entering wait loop'&@lf)
		If $ar[0] <> '' And $ar[0] <> ' ' Then
			While 0 <> 1 
				$type = ControlGetText($h_AcuWin,'',$h_unit_type)
				If $type <> '' And $type <> ' ' Then
					ConsoleWrite('type found, deleting unit'&@lf)
					_SendTab('{ENTER}',1)
					_SendTab('!{D 2}',1)
					_SendTab('{ENTER}',1)
					ExitLoop
				ElseIf WinExists('Unit Maintenance', 'Unit is not on file.') Then
					ConsoleWrite('window found, skipping'&@lf)
					_SendTab('{n}',1)
					ExitLoop
				EndIf
			WEnd
			ConsoleWrite('loop exited, read next'&@lf)
		EndIf
		
		$ar = FileReadline($file)
	Until @error = -1 
	BlockInput(0)
	ConsoleWrite('delete process finished.'&@lf)
	
EndFunc