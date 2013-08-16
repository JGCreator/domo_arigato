Opt("WinTitleMatchMode", 4)	 
; Script Start - Add your code below here


#cs----
Description:
	Change from 1 Acu-tab to another
	
		_ChangeTabs ( "Tab ID", Directrion, Byref Current Tab )

Parameters:
	Tab ID 			= a string value for the control ID
	Direction 		= a single char string to tell the function which way to tab
	Current Tab		= a numeric value to identify where the control started. [numbers might be expected to increase from left to right starting at 1]
	
Return Values:	(fail return code not implemented (07/29/13))
	Success: value of CurrentTab
	Fail: -1
	
Remarks: 
	Future changes may ask the user to enter the max number of tabs to allow cyclical tabbing, meaning that if the 
	calling function sends a right direction and the current tab = max tabs it will perform a left operation until 
	current tab = 1.
	
	Invalid direction entries should give fail return code.
	
	Missing required variables should give fail return code.
	
#ce----
Func _ChangeTabs($TabControl, $Direction, ByRef $CurrentTab)
   ControlClick("", "", $TabControl)
   ControlFocus("", "", $TabControl)
   If $Direction = "R" Then
	  Send("{RIGHT}")
	  $CurrentTab += 1
   Else
	  If $Direction = "L" Then
		 Send("{LEFT}")
		 $CurrentTab -= 1
	  Else
	     ;Not a valid direction
	  EndIf
   EndIf
EndFunc


#cs----
0 = success 
1 = fail
#ce----
Func _OpenAcuProgram($ProgramName)
	Opt("WinTitleMatchMode", 4)	 
	If WinExists ( "Advantage - Dubuque Data", '') Then
		If not WinActive("Advantage - Dubuque Data") then 
			WinActivate("Advantage - Dubuque Data")
		EndIf
		$txtLauncher = "[CLASS:Edit; INSTANCE:1]"
		ControlFocus("", "", $txtLauncher)
		_EnterTheText("","",$txtLauncher, $ProgramName, 4)
		
		$Status = ''
		
		;While not $Status = $ProgramName
			$Status = ControlGetText('','',$txtLauncher)
			If $Status = $ProgramName Then
				send("{ENTER}")
				return 0
			Else
				
				return 1
			EndIf
		;WEnd
	EndIf
EndFunc


#cs----

#ce----
Func _ValueCompare($ControlToCheck, $ValueToCheck)
   $ControlText = ControlGetText("", "", $ControlToCheck)
   If $ControlText = $ValueToCheck Then
	  Return "Equal"
   Else
	  Return "Not Equal"
   EndIf
EndFunc


#cs----

#ce----
Func _ClickTheButton ( $CTB_WinTTL, $CTB_WinTxt, $CTB_BtnID, $CTB_Pause )
	$CTB_Clicked	=	''
	Sleep ( $CTB_Pause * 1000 / 2 )
	While Not WinExists ( $CTB_WinTTL, $CTB_WinTxt )
		Sleep ( $CTB_Pause )
	WEnd

	If WinExists ( $CTB_WinTTL, $CTB_WinTxt ) Then
		While $CTB_Clicked < 1
			If Not WinActive ( $CTB_WinTTL, $CTB_WinTxt ) Then WinActivate ( $CTB_WinTTL, $CTB_WinTxt )
			If ControlFocus ( $CTB_WinTTL, $CTB_WinTxt , $CTB_BtnID) Then
				$CTB_Clicked = ControlClick ( $CTB_WinTTL, $CTB_WinTxt, $CTB_BtnID, 'LEFT', 1 )
			Else
			;unable to get control focus
			EndIf
		WEnd
	EndIf
EndFunc


#cs----
$ETT_WinTTL	= window title
$ETT_WinTxt = window text
$ETT_EditID = control ID to set text
$ETT_Entry 	= text to set 
$ETT_Pause 	= 4 = 1 second
#ce----
Func _EnterTheText ( $ETT_WinTTL, $ETT_WinTxt, $ETT_EditID, $ETT_Entry, $ETT_Pause )
	$ETT_Entered		=	''
	$ETT_EscCount	=	0
	
	; Loop until the window exists
	While Not WinExists ( $ETT_WinTTL, $ETT_WinTxt )
		Sleep ( $ETT_Pause * 1000 / 2 )
		$ETT_EscCount += 1
		If $ETT_EscCount > 50 Then
			MsgBox ( 4096, 'Error', 'Error Waiting for: ' & $ETT_WinTTL & ' - ' & $ETT_WinTxt )
			Return 1
		EndIf
	WEnd
	
	; Reset the counter
	$ETT_EscCount	=	0
	
	; Loop until Text set or get focus abandoned.
	While $ETT_Entered < 1
		If Not WinActive ( $ETT_WinTTL, $ETT_WinTxt ) Then WinActivate ( $ETT_WinTTL, $ETT_WinTxt )
		If ControlFocus ( $ETT_WinTTL, $ETT_WinTxt, $ETT_EditID ) Then
			$ETT_Entered = ControlSetText ( $ETT_WinTTL, $ETT_WinTxt, $ETT_EditID, $ETT_Entry )
		Else
			;unable to get control focus
;~ 			_logger("unable to get control focus on attempt " & $ETT_EscCount)
			$ETT_EscCount += 1
			Sleep($ETT_Pause * 1000 / 2)	
			
			; stop trying after 5 attempts and 5 seconds
			if $ETT_EscCount > 5 Then
;~ 				_logger("Canceling Process After " & $ETT_EscCount & " attempts.")
				Return 1
			EndIf
		EndIf
	WEnd
EndFunc

#cs----

#ce----
Func _SendTab( $String, $Pause )
   Send($String)
   Sleep( $Pause * 1000 / 2)
EndFunc

#cs----

#ce----
Func _WaitActivate( $String, $Pause )
	$naps = 0
	While Not WinExists ($String)
		Sleep ( 500 )
		$naps = $naps + 1
		If $naps > 9 Then
			return 1
		EndIf
	WEnd
	
   WinWaitActive($String)
   Sleep( $Pause * 1000 / 2)
EndFunc

Func _Acu_DropDown($Control, $Value)
   _ClickTheButton("", "", $Control, 1)
EndFunc

Func _Acu_DropList($Control, $Selection)
   _ClickTheButton("", "", $Control, 1)
EndFunc

