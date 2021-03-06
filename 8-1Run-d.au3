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

	Global $Window = GUICreate("8-1Run-d",125,110)	
	
	Global $InstallDir = 'C:\Program Files\Acucorp\Acucbl810\'
	Global $bDebug = ''
	GUICtrlCreateLabel("Server name:", 5,5)
	GUICtrlCreateLabel('Site number:',5,45)
	
	
	
	; add an options menu and start with "run as -d (debug)".
	; set check status of options according to key(s) value.
	; change operation according to key(s) value.
	
	; look for previous server key and suggest as default
	$ReadSuccess = RegRead("HKEY_CURRENT_USER\SOFTWARE\USERDEF\AcuOpen", "LastServer")
	ConsoleWrite('Server key: ' & $ReadSuccess & @lf)
	If $ReadSuccess <> '' Then
		Global $hHost = GUICtrlCreateInput($ReadSuccess,5,20,115,20)
	Else	
		Global $hHost = GUICtrlCreateInput('',5,20,115,20)
	EndIf
	GUICtrlSetTip($hHost, "ex: dds630-01 or dds971-2")
	
	
	; look for previous site key and suggest as default
	$ReadSuccess = RegRead("HKEY_CURRENT_USER\SOFTWARE\USERDEF\AcuOpen", "LastLogin")
	ConsoleWrite('Login key: ' & $ReadSuccess & @lf)
	IF $ReadSuccess <> '' And $ReadSuccess <> 0 Then
		Global $hSite = GUICtrlCreateInput($ReadSuccess,5,60,50,20)
	Else
		Global $hSite = GUICtrlCreateInput('',5,60,50,20)
	EndIf
	GUICtrlSetTip($hSite, "Optional to login as admin")
	
	
	; look for previous debug setting key and use last value
	Global $hDebug = GUICtrlCreateCheckbox('debug', 65,60)
	$ReadSuccess = RegRead("HKEY_CURRENT_USER\SOFTWARE\USERDEF\AcuOpen", "DebugBool")
	ConsoleWrite('Debug key: ' & $ReadSuccess & @lf)
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






Func _okayClicked()
	ConsoleWrite('okay clicked or enter presed' & @LF)
	
	$sName = ControlGetText('','',$hHost)
	$nSite = ControlGetText('','',$hSite)
	
	$Split = StringSplit($sName, "-")
	If $Split[0] < 2 Then 	; no delimiter found
		ConsoleWrite('Error:' & @lf & chrw(9) & 'Invalid parameter. Unable to identify the server by the value given. (ex: dds630-1)' & @lf)
		MsgBox(0, 'Error:', 'Invalid parameter. Unable to identify the server and/or site by the value given. (required delimiter = "-")' & @lf)
;~ 		Exit
		return 1
	EndIf
	
	If Not StringIsDigit($Split[2]) Then
		ConsoleWrite('Error:' & @lf & chrw(9) & 'Parameter is not number' & @lf)
		MsgBox(0, 'Error:', 'Parameter is not number' & @lf)
;~ 		Exit
		return 1
	EndIf
	
	; make 2 digit instance if needed
	if stringlen($Split[2]) < 2 Then
		$Split[2] = '0'&$Split[2]
	EndIf
	
	; write the check passing value to registry
	$Write = RegWrite("HKEY_CURRENT_USER\SOFTWARE\USERDEF\AcuOpen", "LastServer", "REG_SZ", $sName)
	If $Write = 0 Then 
		ConsoleWrite('Error:' & @lf & chrw(9) & 'An error occurred trying to open and write the server key.' & @lf)
		MsgBox(0, 'Error:', "Can't open and write key." & @lf)
;~ 		Exit
	EndIf
	
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
	
	
	; string box name for call
	$sName = $Split[1] & 'vmub' & $Split[2]
	; define var for -d option
	$option = GUICtrlRead($hDebug)
	
	; Before calling the run command, search for the Acucorp\Acucbl810 directory in {C:\Program Files\} 
	; If none is found then prompt the user with a message requiring an Acu8.1 installation.
	If FileExists($InstallDir) _ 
		And StringInStr(filegetattrib($InstallDir),"D") Then
			
			; evaluate the checkbox and write reg key
			ConsoleWrite('checking debug' & @lf)			
			If $option = 1 Then
				$option = ' -d '
				$Write = RegWrite("HKEY_CURRENT_USER\SOFTWARE\USERDEF\AcuOpen", "DebugBool", "REG_SZ", 'True')
				If $Write = 0 Then 
					ConsoleWrite('Error:' & @lf & chrw(9) &'An error occurred trying to open and write the debug key.' & @lf)
					MsgBox(0, 'Error:' , "Can't open and write key." & @lf)
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
			EndIf
	Else
		MsgBox(0,'Acu 8.1 Not Found','This script requires an installation of Acucbl810 in C:\Program Files\Acucorp')
		ConsoleWrite('Error:' & @lf & chrw(9) & 'Installation not found' & _ 
						@LF & chrw(9) & 'This script requires an installation of Acucbl810 in C:\Program Files\Acucorp' & @LF)
		Exit
	EndIf
			

	; look for connection fail while checking for debug / Advantage window
	ConsoleWrite('enter connection loop' & @lf)

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
		WEnd
	
	; activate and click the go button if debug used
	If $option = ' -d ' Then
		WinActivate("ACUCOBOL-GT Debugger")
		ConsoleWrite('clicking the button' & @lf)
		_ClickTheButton ( "ACUCOBOL-GT Debugger", '', "[CLASS:AcuBitButtonClass; INSTANCE:11]", .25)
	EndIf
	
	; if login parameter is true, look for the window and call the login function
	If $login = True Then
		WinWait('Advantage Login')
		ConsoleWrite('call login function' & @LF)
		_Advantage_Login("Admin", "United", $nSite)
		Return 1
	EndIf
	
	; put focus back into the server box.
	ControlFocus('','',$hHost)
EndFunc
	
Func _cancelClicked()
	ConsoleWrite('exiting' & @lf)
	Exit
EndFunc
	