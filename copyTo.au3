#include <GUIConstantsEx.au3>
opt('GUIOnEventMode',1)
;~ $Source = 'C:\Acucorp\workdir\jgust_view\adv1020\invoicer\OBJECT\frmInvoicer.acu'
;~ $Dest = '\\emp267vmub01\u\AGO\'



;~ ConsoleWrite('result is: ' & $result & @lf)
_cpTo()

Func _cpTo()

	
	
	$sTitle = InputBox("Name session?", 'If you want to name this session, ' & @lf & 'enter the name in the box.' _
							  & @lf & '(short names recommended)')
	
	If $sTitle <> '' Then
		$Window = GUICreate($sTitle,200,130)
	Else	
		$Window = GUICreate("Copy to...",200,130)
	EndIf
	
	GUICtrlCreateLabel("Source:", 5,5)
	GUICtrlCreateLabel('Destination:',5,45)
	
	
	Local $filemenu, $FileOpen
	$filemenu = GUICtrlCreateMenu("File")
	$FileSource = GUICtrlCreateMenuItem("Source...", $filemenu)
	$FileDest = GUICtrlCreateMenuItem("Dest...", $filemenu)
	
	; start the dialog from reg value saved
	$Sourcepath = RegRead("HKEY_CURRENT_USER\SOFTWARE\USERDEF\CPto", "LastSourceDir")
	if $Sourcepath <> '' then
		Global $hSource = GUICtrlCreateInput($Sourcepath ,5,20,190,20)
	Else
		Global $hSource = GUICtrlCreateInput('',5,20,190,20)
	EndIf
	GUICtrlSetTip($hSource, "Soruce for copy paste")


	; start the dialog from reg value saved
	$Destpath = RegRead("HKEY_CURRENT_USER\SOFTWARE\USERDEF\CPto", "LastDestDir")
	if $Destpath <> '' then
		Global $hDest = GUICtrlCreateInput($Destpath,5,60,190,20)
	Else
		Global $hDest = GUICtrlCreateInput('',5,60,190,20)
	EndIf
	GUICtrlSetTip($hDest, "Destination for copy paste")
	
	
	$okbutton = GUICtrlCreateButton("Go", 50,85,100,20)
;~ 	Global $cancelbutton = GUICtrlCreateButton("Cancel",65,85,50,20)
	
	$hENTER = GUICtrlCreateDummy()
	$hEsc = GUICtrlCreateDummy()
	dim $aAccel[2][2] 
	$aAccel[0][0] = '{ENTER}' 
	$aAccel[0][1] = $hENTER
	$aAccel[1][0] = '{Esc}'
	$aAccel[1][1] = $hEsc
	GUISetAccelerators($aAccel)
	
	; set events for user interaction.
	GUICtrlSetOnEvent($hENTER, "_GoClicked")
	GUICtrlSetOnEvent($okbutton, "_GoClicked")
	GUICtrlSetOnEvent($filesource, "_openSource")
	GUICtrlSetOnEvent($filedest, "_openDest")
	
	GUICtrlSetOnEvent($hEsc, '_exit')
	GUISetOnEvent($GUI_EVENT_CLOSE, "_exit")
	GUISetState(@SW_SHOW, $window)
	ControlFocus('','',$hSource)
	
	; wait for user event
	while 1
		sleep(1000)
	WEnd
	
	
EndFunc

Func _exit()
	Exit
EndFunc

Func _openSource()
	$Sourcepath = RegRead("HKEY_CURRENT_USER\SOFTWARE\USERDEF\CPto", "LastSourceDir")
	ConsoleWrite($sourcepath & @lf)
	
	; open the dialog
	If $Sourcepath <> '' Then
		$Sourcefile = FileOpenDialog("Choose file...", $Sourcepath, "All (*.*)", 10)
	Else
		$Sourcefile = FileOpenDialog("Choose file...", 'C:\', "All (*.*)", 10)
	EndIf
				
	; when successful
	ControlSetText ( "", "", $hSource, $Sourcefile )
	$sourcefile = ''
	
EndFunc

Func _openDest()
	$Destpath = RegRead("HKEY_CURRENT_USER\SOFTWARE\USERDEF\CPto", "LastDestDir")
	ConsoleWrite($Destpath & @lf)			
	; open the dialog
	If $Destpath <> '' Then
		$Destfile = FileOpenDialog("Choose file...", $Destpath, "All (*.*)", 10)
	Else
		$Destfile = FileOpenDialog("Choose file...", 'C:\', "All (*.*)", 10)
	EndIf
				
	; when successful
	ControlSetText ( "", "", $hDest, $Destfile )
	$destfile = ''
EndFunc

Func _GoClicked()
	$stat = GUICtrlCreateLabel('',5,85,35,20)
	$source = ControlGetText('','',$hSource)
	$dest = ControlGetText('','',$hDest)
	RegWrite("HKEY_CURRENT_USER\SOFTWARE\USERDEF\CPto", "LastSourceDir", "REG_SZ", $source)
	RegWrite("HKEY_CURRENT_USER\SOFTWARE\USERDEF\CPto", "LastDestDir", "REG_SZ", $dest)
	
	$result = FileCopy($Source, $dest, 1)
	 
	If $result = 1 Then
		GUICtrlSetData($stat, "Good")
		sleep(1000)
		GUICtrlSetData($stat, '')
	Else
		GUICtrlSetData($stat, "Bad")
		sleep(1000)
		GUICtrlSetData($stat, '')
	EndIf

EndFunc


