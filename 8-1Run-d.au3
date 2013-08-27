#cs----
Description:
	Run an instance of advantage with the debug option.
	
Remarks:
	This program assumes that the users environment has an installation of Acu 8.1.x in 'c:\Program Files\Acucourp\...'
	When using for the first time, the message box appears empty, but every time the box prompts for the last used entry.
	This last used entry is written to registry under HKEY_CURRENT_USER\SOFTWARE\USERDEF\LastLogin

#ce----

#include "./Includes/CommonControlFunctions.au3"

$ReadSuccess = RegRead("HKEY_CURRENT_USER\SOFTWARE\USERDEF\AcuOpen", "LastLogin")

If $ReadSuccess <> '' Then
	$BoxName = InputBox('Short Box Name', 'Enter the server to run against. (ex. 630-01)', _ 
						$ReadSuccess,'', 150,130,0,892)
Else	
	$BoxName = InputBox('Short Box Name', 'Enter the server to run against.' & @lf & "(ex. 630-01)")
EndIf

if $BoxName = '' then Exit
	
$Split = StringSplit($BoxName, "-")
If $Split[0] <> 2 Then
	ConsoleWrite('Invalid parameter. The delimiter "+" is needed between the 3rd and 4th digit.')
	Exit
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
	_ClickTheButton ( "ACUCOBOL-GT Debugger", '', "[CLASS:AcuBitButtonClass; INSTANCE:11]", 1)
	; call the login function.
EndIf

Exit
	

