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

#include "./Includes/CommonControlFunctions.au3"
#include "./Includes/Advantage_login.au3"

opt('WinTextMatchMode', 2)

$Result = 1

; loop function call until success or critical error.
;~ While $Result <> 2 And $Result <> 0
While $Result <> 0
	$Result = _Run8_1()
WEnd

Exit




Func _Run8_1()
	; look for previous success key and suggest as default
	$ReadSuccess = RegRead("HKEY_CURRENT_USER\SOFTWARE\USERDEF\AcuOpen", "LastLogin")

	If $ReadSuccess <> '' Then
		$BoxName = InputBox('Short Box Name', 'Enter the server and site (optional) to run against.' & @lf & _
							'(ex. 630-01-01)', $ReadSuccess,'', 150,150)
	Else	
		$BoxName = InputBox('Short Box Name', 'Enter the server and site (optional) to run against.' & @lf & _ 
							"(ex. 630-01-01)",'','',150,150)
	EndIf


	; error catching on parameters
	
	; return success if user cancels
	if $BoxName = '' then return 0
		
	$Split = StringSplit($BoxName, "-")
	If $Split[0] < 2 Then 	; no delimiter found
		ConsoleWrite('Invalid parameter. Unable to identify the box and/or site by the value given. (required delimiter = "-")' & @lf)
		MsgBox(0, 'Invalid parameter. Unable to identify the box and/or site by the value given. (required delimiter = "-")' & @lf)
;~ 		Exit
		return 1
	EndIf

	If Not StringIsDigit($Split[1]) Or Not StringIsDigit($Split[2]) Then
		ConsoleWrite('Error: Parameter is not number' & @lf)
		MsgBox(0, 'Error:', 'Parameter is not number' & @lf)
;~ 		Exit
		return 1
	EndIf

	If $Split[0] = 3 And Not StringIsDigit($Split[3]) Then
		ConsoleWrite('Error: Parameter is not number')
		MsgBox(0, 'Error:', 'Parameter is not number')
;~ 		Exit
		return 1
	EndIf


	; make 2 digit instance if needed
	if stringlen($Split[2]) < 2 Then
		$Split[2] = '0'&$Split[2]
	EndIf

	; write the check passing value to registry
	$Write = RegWrite("HKEY_CURRENT_USER\SOFTWARE\USERDEF\AcuOpen", "LastLogin", "REG_SZ", $BoxName)
	If $Write = 0 Then 
		ConsoleWrite('An error occurred trying to open and write the key.' & @lf)
		MsgBox(0, 'Error:', "Can't open and write key." & @lf)
;~ 		Exit
	EndIf

	; string box name for call
	$BoxName = 'dds' & $Split[1] & 'vmub' & $Split[2]

	; call run function and check pid
	$PID = run('"C:\Program Files\Acucorp\Acucbl810\AcuGT\bin\acuthin.exe" --nosplash --password nfsuser ' & $BoxName & ' -d menu')
	If $PID = 0 and @error <> 0 Then	
		MsgBox(0, "Unable to Launch Process", "Check that the box is valid and running, " & @lf & _
		"and that the installation is located @:" & @lf & '"C:\Program Files\Acucorp\Acucbl810\AcuGT\bin\acuthin.exe"'
		 ConsoleWrite("Unable to Launch Process. Check that the box is valid and running, " & @lf & _
		"and that the installation is located @:" & @lf & '"C:\Program Files\Acucorp\Acucbl810\AcuGT\bin\acuthin.exe"'& @lf)
	Else
		If WinExists("ACUCOBOL-GT Debugger") Then WinActivate("ACUCOBOL-GT Debugger")
		_ClickTheButton ( "ACUCOBOL-GT Debugger", '', "[CLASS:AcuBitButtonClass; INSTANCE:11]", .25)
	EndIf

	; look for advantage login window while checking for failed connection
	While not WinExists("Advantage Login")
		
		If WinExists('ACUCOBOL-GT Thin Client', 'Connection failed') Then
			WinWaitClose ( "ACUCOBOL-GT Thin Client", "Connection failed")
			Return 1
		EndIf

	WEnd

	; login window found, call the login function if parameter specified
	If $Split[0] = 3 Then
		_Advantage_Login("Admin", "United", $Split[3])
		Return 0
	EndIf

EndFunc
	