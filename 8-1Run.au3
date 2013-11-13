#cs----
Description:
	Run an instance of advantage with the debug option.
	
Parameters:
	Input box value must represent a networked virtual server running acu8.1.x [site number for login function is optional]
	
Remarks:
	This program assumes that the users environment has an installation of Acu 8.1.x in 'c:\Program Files\Acucourp\...'
	
	When using for the first time, the message box appears empty, but every time the box prompts for the last used entry.
	This last used entry is written to registry under HKEY_CURRENT_USER\SOFTWARE\USERDEF\LastLogin
	
	The format of the input box string requires a '-' (dash) delimiter between the box id, the box instance[, and the site #].
	{boxID}  {Instance} [{site #}]
	  630   - 	  01   -    01		= 630-01-01  ---> will run on dds630vmub01 site 01 
	  
	If supplying a site # for login function, "Admin" and "United" are assumed for login credentials.

#ce----
#include <GUIConstantsEx.au3>
#include <GuiComboBox.au3>
#include <Array.au3>
;~ #include "./Includes/Downloads/_GetIntersection.au3"
#include "./Includes/CommonControlFunctions.au3"
#include "./Includes/Advantage_login.au3"

opt('WinTextMatchMode', 2)
opt('WinTitleMatchMode',1)
opt('TrayIconDebug', 1)
opt('GUIOnEventMode',1)

$Result = 1

; loop function call until success or critical error.
;~ While $Result <> 2 And $Result <> 0
;~ While $Result <> 0
;~ 	ConsoleWrite('calling _Run8_1()' & @LF)
;~ 	$Result = _Run8_1()
;~ WEnd

;~ Exit

_Run8_1()


Func _Run8_1()

	Global $Window = GUICreate("Run Adv",125,110)
	
	Global $WinList1
	Global $WinList2
;~ 	$WinList1[0][0] = 0
;~ 	$WinList2[0][0] = 0
	Global $arResults
	
	Global $InstallDir = 'C:\Program Files\Acucorp\Acucbl810\'
	Global $bDebug = ''
	GUICtrlCreateLabel("Server name:", 5,5)
	GUICtrlCreateLabel('Site number:',5,45)
	
	
	
	; add an options menu and start with "run as -d (debug)".
	; set check status of options according to key(s) value.
	; change operation according to key(s) value.
	

   ; check for ...\Program Files\8-1Run\serverList.txt exists
   If FileExists(@ProgramFilesDir & "\8-1Run\serverList.txt") Then
	  ; true, use file to populate the combobox
	  Global $hSvrFile = FileOpen(@ProgramFilesDir & "\8-1Run\serverList.txt")
	  $txt = ''
	  While 1
		 $txt &= FileReadLine($hSvrFile) & "|"
		 If @error = -1 Then 
			$txt = StringTrimRight($txt,2)
			ExitLoop
		 EndIf
	  WEnd	  
	  Global $hHost = _GUICtrlComboBox_Create($Window,$txt,5,20,115,20)
   Else
	  ; false, assume the dir doesn't exist either and create both
	  DirCreate(@ProgramFilesDir & "\8-1Run\")
	  Global $hSvrFile = FileOpen(@ProgramFilesDir & "\8-1Run\serverList.txt",1)
	  Global $hHost = _GUICtrlComboBox_Create($Window,'',5,20,115,20)
   EndIf
   FileClose($hSvrFile)
   
	; look for previous server key and suggest as default
	$ReadSuccess = RegRead("HKEY_CURRENT_USER\SOFTWARE\USERDEF\AcuOpen", "LastServer")
	ConsoleWrite('Server key: ' & chrw(9) & $ReadSuccess & @lf)
   If $ReadSuccess <> '' Then
	  _GUICtrlComboBox_SetEditText($hHost,$ReadSuccess)		
   EndIf
   
	
	; look for previous site key and suggest as default
	$ReadSuccess = RegRead("HKEY_CURRENT_USER\SOFTWARE\USERDEF\AcuOpen", "LastLogin")
	ConsoleWrite('Login key: ' & chrw(9) & $ReadSuccess & @lf)
	IF $ReadSuccess <> '' And $ReadSuccess <> 0 Then
		Global $hSite = GUICtrlCreateInput($ReadSuccess,5,60,50,20)
	Else
		Global $hSite = GUICtrlCreateInput('',5,60,50,20)
	EndIf
	GUICtrlSetTip($hSite, "Optional to login as admin")
	
	
	; look for previous debug setting key and use last value
	Global $hDebug = GUICtrlCreateCheckbox('debug', 65,60)
	$ReadSuccess = RegRead("HKEY_CURRENT_USER\SOFTWARE\USERDEF\AcuOpen", "DebugBool")
	ConsoleWrite('Debug key: ' & chrw(9) & $ReadSuccess & @lf)
	If $ReadSuccess <> '' Then 
		Select
		case $ReadSuccess = 'True'
			GUICtrlSetState($hDebug, $GUI_CHECKED)
		case Else
			GUICtrlSetState($hDebug, $GUI_UNCHECKED)
		EndSelect
	EndIf
	
	
	Global $okbutton = GUICtrlCreateButton("OK", 10,85,50,20)
	Global $cancelbutton = GUICtrlCreateButton("Cancel",65,85,50,20)
	
	Global $hENTER = GUICtrlCreateDummy()
	Global $hEsc = GUICtrlCreateDummy()
	dim $aAccel[2][2] 
	$aAccel[0][0] = '{ENTER}' 
	$aAccel[0][1] = $hENTER
	$aAccel[1][0] = '{Esc}'
	$aAccel[1][1] = $hEsc
	GUISetAccelerators($aAccel)
	
	; set events for user interaction.
	GUICtrlSetOnEvent($hENTER, "_okayClicked")
	GUICtrlSetOnEvent($okbutton, "_okayClicked")
	GUICtrlSetOnEvent($hEsc, '_cancelClicked')
	GUICtrlSetOnEvent($cancelbutton, "_cancelClicked")
	GUISetOnEvent($GUI_EVENT_CLOSE, "_cancelClicked")
	GUISetState(@SW_SHOW, $window)
	ControlFocus('','',$hHost)
	
	; wait for user event
	while 1
		sleep(1000)
	WEnd
	
	
EndFunc





;==================================================================================================
; Function Name:   	_okayClicked()
; Description::    	event for okay click / enter press
;						- validate parameters
;						- open new instance of advantage 
;
; Parameter(s):    	read from screen controls
;
; Return Value(s): 	1 - an error occurred - return to main window 
; Note:            	- run with command prompt to produce consol logging
; Author(s):       	josh gust (ddslive.com)
;==================================================================================================
Func _okayClicked()
	ConsoleWrite('okay clicked or enter presed' & @LF)
	$sName = ControlGetText('','',$hHost)
	$nSite = ControlGetText('','',$hSite)
	
	; validate site parameter and set boolean value
	$login = True
	If $nSite = '' Then
		$login = False
	EndIf
	
	$nSite = Number($nSite)
	If $nSite = 0 Or Not IsInt($nSite) Then
		$login = False
	Else
		$Write = RegWrite("HKEY_CURRENT_USER\SOFTWARE\USERDEF\AcuOpen", "LastLogin", "REG_SZ", $nSite)
		If $Write = 0 Then 
			ConsoleWrite('Error:' & @lf & chrw(9) &'An error occurred trying to open and write the site key.' & @lf)
			MsgBox(0, 'Error:' , "Can't open and write key." & @lf)
		EndIf
	EndIf
	
	; define var for -d option
	$option = GUICtrlRead($hDebug)
	
	; Before calling the run command, search for the Acucorp\Acucbl810 directory in {C:\Program Files\} 
	; If none is found then prompt the user with a message requiring an Acu8.1 installation.
	If FileExists($InstallDir) _ 
		And StringInStr(filegetattrib($InstallDir),"D") Then
			
			; evaluate the checkbox and write reg key
			ConsoleWrite('checking debug box' & @lf)			
			If $option = 1 Then
				$option = ' -d '
				$Write = RegWrite("HKEY_CURRENT_USER\SOFTWARE\USERDEF\AcuOpen", "DebugBool", "REG_SZ", 'True')
				If $Write = 0 Then 
					ConsoleWrite('Error:' & @lf & chrw(9) &'An error occurred trying to open and write the debug key.' & @lf)
					MsgBox(0, 'Error:' , "Can't open and write key." & @lf)
				EndIf

				; get a 2d list of debug windows
				$WinList1 = WinList('ACUCOBOL-GT Debugger')
				ConsoleWrite('WinList1 count: ' & $WinList1[0][0] & @LF)
				If $WinList1[0][0] = 0 then 
					$first = True
				Else 
					$first = False
;~ 					$WinList2[0][0] = $WinList1[0][0]
					ConsoleWrite('getting handle(s) from WinList 1' & @LF)
					$WinList1 = _ListDebugHandles($WinList1)
				EndIf
			Else
				$option = ' '
				$Write = RegWrite("HKEY_CURRENT_USER\SOFTWARE\USERDEF\AcuOpen", "DebugBool", "REG_SZ", 'False')
				If $Write = 0 Then 
					ConsoleWrite('Error:' & @lf & chrw(9) &'An error occurred trying to open and write the debug key.' & @lf)
					MsgBox(0, 'Error:' , "Can't open and write key." & @lf)
				EndIf
			EndIf
			
			; call run function and check pid
			ConsoleWrite('calling run()' & @LF)
			$PID = run('"C:\Program Files\Acucorp\Acucbl810\AcuGT\bin\acuthin.exe" ' & _ 
						'--nosplash --password nfsuser ' & $sName & $option & 'menu')
			If $PID = 0 and @error <> 0 Then	
				MsgBox(0, "Unable to Launch Process", "Check that the box is valid and running, " & @lf & _
				"and that the installation is located @:" & @lf & '"C:\Program Files\Acucorp\Acucbl810\AcuGT\bin\acuthin.exe"'
				ConsoleWrite('Error:' & @lf & chrw(9) & _ 
							"Unable to Launch Process. Check that the box is valid and running, " & @lf & chrw(9) & _
							"and that the installation is located @:" & @lf & chrw(9) & - 
							'"C:\Program Files\Acucorp\Acucbl810\AcuGT\bin\acuthin.exe"'& @lf)
			Else
			   ; write the check passing value to registry | only write after successful run
			   $Write = RegWrite("HKEY_CURRENT_USER\SOFTWARE\USERDEF\AcuOpen", "LastServer", "REG_SZ", $sName)
			   If $Write = 0 Then 
				  ConsoleWrite('Error:' & @lf & chrw(9) & 'An error occurred trying to open and write the server key.' & @lf)
				  MsgBox(0, 'Error:', "Can't open and write key." & @lf)
			   EndIf

			   ; check the server file and write new line if needed
			   ; ...
			EndIf
	Else
		MsgBox(0,'Acu 8.1 Not Found','This script requires an installation of Acucbl810 in C:\Program Files\Acucorp')
		ConsoleWrite('Error:' & @lf & chrw(9) & 'Installation not found' & _ 
						@LF & chrw(9) & 'This script requires an installation of Acucbl810 in C:\Program Files\Acucorp' & @LF)
		Exit
	EndIf
			

	; look for connection fail while checking for debug / Advantage window
	ConsoleWrite('enter connection loop' & @lf)
		dim $trycount =0
		While not WinExists("ACUCOBOL-GT Debugger") _ 
					And Not WinExists("Advantage - ")
		
			If WinExists('ACUCOBOL-GT Thin Client', 'Connection failed') Then 
				$wintxt = WinGetText('ACUCOBOL-GT Thin Client')
				$wintxt = WinGetText('Error')
				$trimcount = StringInStr($wintxt, @lf)
				$wintxt = StringTrimLeft($wintxt, $trimcount)
				ConsoleWrite('Error:' & @lf & chrw(9) & $wintxt & @lf)
				WinWaitClose ( "ACUCOBOL-GT Thin Client", "Connection failed")
				ConsoleWrite('window closed' & @lf)
				Return 1
			EndIf
		
			If WinExists('Error', 'Program missing or inaccessible') Then 
				$wintxt = WinGetText('Error')
				$trimcount = StringInStr($wintxt, @lf)
				$wintxt = StringTrimLeft($wintxt, $trimcount)
				ConsoleWrite('Error:' & @lf & chrw(9) & $wintxt & @lf)
				WinWaitClose ('Error', 'Program missing or inaccessible')
				ConsoleWrite('window closed' & @lf)
				Return 1
			EndIf
			
			$trycount +=1			
			if $trycount > 60 then 
				If MsgBox(4, 'Time out', 'Do you want to keep waiting?') = 6 Then
					$trycount = 0
				else 
					Return 1
				EndIf
			EndIf
			sleep( 250 )	
		WEnd
	
	; activate and click the go button if debug used
	dim $hWin
	
	If $option = ' -d ' Then
		If $first = False Then
			; get a list of windows after the new opens
			Do 
				Sleep( 250 )
				$WinList2 = WinList("ACUCOBOL-GT Debugger")
			Until UBound($WinList2, 1) > UBound($WinList1)
			
			; chop the names off the 2d array to only give 1d array and handles
			ConsoleWrite('WinList2 count: ' & $WinList2[0][0] & @LF)
			ConsoleWrite('getting handle(s) from WinList 2' & @LF)
			$WinList2 = _ListDebugHandles($WinList2)
			
			; compare list before run() to list after to find new window
			For $i =0 To UBound($WinList2)-1
				_ArraySearch($WinList1, $WinList2[$i])
				If @error = 6 Then	; value not found in array, use this one	
					ConsoleWrite('@error: ' & @error & ' array diff found' & @LF)
					$hWin = $WinList2[$i]
					ConsoleWrite('hWin: ' & $hWin & @LF)
					$i = UBound($WinList2)-1
					ExitLoop	; if this condition is met, then the until on the outer loop must also be met. So, this exits both loops.
				Else
					ConsoleWrite($WinList2[$i] & ' : not new' & @lf)
				EndIf
			Next
		EndIf			
		
		ConsoleWrite('clicking $hWin' & @lf)
		WinActivate($hWin)	
		ConsoleWrite('clicking the button' & @lf)
		_ClickTheButton ( $hWin, '', "[CLASS:AcuBitButtonClass; INSTANCE:11]", .25)
	EndIf
	
	; if login parameter is true, look for the window and call the login function
	If $login = True Then
		WinWait('Advantage Login','',1)
		ConsoleWrite('call login function' & @LF)
		_Advantage_Login("Admin", "United", $nSite)
	EndIf
	
	; put focus back into the server box.
	ControlFocus('','',$hHost)
	ConsoleWrite(@lf)
	Return 1
EndFunc

;==================================================================================================
; Function Name:   _cancelClicked()
; Description::    exit the program when close event
; Parameter(s):    none
; Return Value(s): Succes	close the script
; Note:            
; Author(s):       josh gust (ddslive.com)
;==================================================================================================
Func _cancelClicked()
	ConsoleWrite('exiting' & @lf)
	Exit
EndFunc

;==================================================================================================
; Function Name:   _ListDebugHandles($set)
; Description::    Get the 1d array of handles from a 2d array of titles + handles
;                  
; Parameter(s):    $Set	(2D-array of title + handle returned from WinList())
;             	   
; Return Value(s): Succes	1D-array    $Return[$handle]
;
; Note:            Comparison is case-sensitiv! - i.e. Number 9 is different to string '9'!
; Author(s):       josh gust (ddslive.com)
;==================================================================================================
Func _ListDebugHandles($set)
	dim $aReturn[ubound($set,1)]
;~ 	ConsoleWrite('ubound aReturn: ' & ubound($aReturn) & @LF)
	For $i = 1 To UBound($set) -1
;~ 		ConsoleWrite('i : ' & $i & @lf)
		ConsoleWrite('$set '&$i&',1: ' & $set[$i][1] & @lf)
		$aReturn[$i-1] = $set[$i][1]
		
	Next
	
	Return $aReturn		
EndFunc	

