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

$ReadSuccess = RegRead("HKEY_CURRENT_USER\SOFTWARE\USERDEF\AcuOpen", "LastLogin")

If $ReadSuccess <> '' Then
	$BoxName = InputBox('Short Box Name', 'Enter the server and site (optional) to run against.' & @lf & '(ex. 630-01-01)', $ReadSuccess,'', 150,150)
Else	
	$BoxName = InputBox('Short Box Name', 'Enter the server and site (optional) to run against.' & @lf & "(ex. 630-01-01)",'','',150,150)
EndIf

if $BoxName = '' then Exit
	
$Split = StringSplit($BoxName, "-")
If $Split[0] < 2 Then
	ConsoleWrite('Invalid parameter. Unable to identify the box and/or site by the value given. (required delimiter = "-")')
	Exit
EndIf

if stringlen($Split[2]) < 2 Then
	$Split[2] = '0'&$Split[2]
EndIf

$Write = RegWrite("HKEY_CURRENT_USER\SOFTWARE\USERDEF\AcuOpen", "LastLogin", "REG_SZ", $BoxName)
If $Write = 0 Then 
	ConsoleWrite('An error occurred trying to open and write the key.')
	Exit
EndIf

$BoxName = 'dds' & $Split[1] & 'vmub' & $Split[2]

$PID = run('"C:\Program Files\Acucorp\Acucbl810\AcuGT\bin\acuthin.exe" --nosplash --password nfsuser ' & $BoxName & ' -d menu')
If $PID = 0 and @error <> 0 Then	
	MsgBox(0, "Unable to Launch Process", "Check that the box is valid and running, " & @lf & _
	"and that the installation is located @:" & @lf & '"C:\Program Files\Acucorp\Acucbl810\AcuGT\bin\acuthin.exe"'
Else
	If WinExists("ACUCOBOL-GT Debugger") Then WinActivate("ACUCOBOL-GT Debugger")
	_ClickTheButton ( "ACUCOBOL-GT Debugger", '', "[CLASS:AcuBitButtonClass; INSTANCE:11]", .25)
EndIf

If $Split[0] = 3 Then
	_Advantage_Login("Admin", "United", $Split[3])
EndIf

Exit
	

