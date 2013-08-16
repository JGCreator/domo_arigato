; Built in libraries
#include <Array.au3>

; Common user includes
#include ".\Includes\CommonControlFunctions.au3"
#include ".\Includes\ODBC_Acu_Query.au3"
#include ".\Includes\SimpleLogger.au3"

; Local includes
#include ".\Includes\FordDealerInfo\FordDealerInfo_Screen.au3"
#include ".\Includes\FordDealerInfo\FordDealerInfo_FieldNames.au3"


; call the main function
;_FordDealerInfo_Main()

; this main function is for testing the Ford Dealer Info program in Advantage
Func _FordDealerInfo_Main()

	; open a logging file 
	_CreateLog("frmFordDealerInfo", 2)
	
;~ Global $file = FileOpen("test.txt", 2)
;~ If $file = -1 Then
;~     MsgBox(0, "Error", "Unable to open file.")
;~     Exit
;~ EndIf

	Opt("WinTitleMatchMode", 4)	 
	Opt("SendKeyDelay", 75)
	Opt("TrayIconDebug", 1) ;1 = on 0 = off



	; Set the value of current tab when the program opens	
	$CurrentTab = 1

	;$FieldArray[

	

	$DSN = InputBox("FordDealerInfo Testing", "I need an AcuODBC DSN for the PC you are testing on.", "")
	
	$ValueToCheck = InputBox("FordDealerInfo Testing", "I need to know what site your are testing.", 1)
	
	$Acu_FileName = "forddlr"
	$ConditionField = "FD_SITE_NO"
	$ConditionOperator = "="


	_logger('waiting for active window')
	_WaitActivate("Advantage - Dubuque Data", 4)

	_logger('opening Ford Dealer Info')
	
	$OpenWait = 0
	Do
		_OpenAcuProgram("frmFordDealerInfo")
		If _WaitActivate("Ford Dealer Information", 4) = 1 Then
			_logger("Program Did Not Open. Check Employee Security.")
			If $OpenWait <= 3 Then
				$OpenWait += 1
				_logger("Open Program Attempt " $OpenWait)
			Else
				_logger("Program Did Not Open. Check Employee Security.")
				_OpenAcuProgram(' ')
				Exit
			EndIf
		Else
			_logger("Program Opened Successfully.")
			$OpenWait = 4
		EndIf
	until $OpenWait > 3
	


	; Validate the current screen controls against the FordDLR File.
	_logger("Checking population of screen to data in file.")
	_VerifyData($DSN, $Acu_Filename, $PACode, $ConditionField, $ConditionOperator, $ValueToCheck, $txtPACode)
	_VerifyData($DSN, $Acu_Filename, $PACheck, $ConditionField, $ConditionOperator, $ValueToCheck, $txtPACheck)
	_VerifyData($DSN, $Acu_Filename, $SubCode, $ConditionField, $ConditionOperator, $ValueToCheck, $txtSubDealer)
	_VerifyData($DSN, $Acu_Filename, $Currency, $ConditionField, $ConditionOperator, $ValueToCheck, $txtCurrency)

	DIM $Fields[3]
	$Fields[0] = $Dist12
	$Fields[1] = $DistAlpha
	$Fields[2] = $Dist456
	_VerifyDistribution($DSN, $Acu_Filename, $Fields, $ConditionField, $ConditionOperator, $ValueToCheck, $txtFordDistributionCode)

	$Fields[0] = $LMDist12
	$Fields[1] = $LMDistAlpha
	$Fields[2] = $LMDist456
	_VerifyDistribution($DSN, $Acu_Filename, $Fields, $ConditionField, $ConditionOperator, $ValueToCheck, $txtLMDistributionCode)

	_VerifyData($DSN, $Acu_Filename, $GEOSALES, $ConditionField, $ConditionOperator, $ValueToCheck, $txtSalesArea)
	_VerifyData($DSN, $Acu_Filename, $Franchise, $ConditionField, $ConditionOperator, $ValueToCheck, $cboFranchiseCode)
	_VerifyData($DSN, $Acu_Filename, $CurrentLabor, $ConditionField, $ConditionOperator, $ValueToCheck, $txtCurrentLaborRate)
	_VerifyData($DSN, $Acu_Filename, $PriorLabor, $ConditionField, $ConditionOperator, $ValueToCheck, $txtPriorLaborRate)
	_VerifyData($DSN, $Acu_Filename, $EffectiveDate, $ConditionField, $ConditionOperator, $ValueToCheck, $dtEffectiveLaborDate)
	_VerifyData($DSN, $Acu_Filename, $DivisionID, $ConditionField, $ConditionOperator, $ValueToCheck, $cboDivisionID)
	_VerifyData($DSN, $Acu_Filename, $DealerCode, $ConditionField, $ConditionOperator, $ValueToCheck, $txtDealerCode)
	_VerifyData($DSN, $Acu_Filename, $Name, $ConditionField, $ConditionOperator, $ValueToCheck, $txtDealerName)
	_VerifyData($DSN, $Acu_Filename, $Address, $ConditionField, $ConditionOperator, $ValueToCheck, $txtAddress)
	_VerifyData($DSN, $Acu_Filename, $City, $ConditionField, $ConditionOperator, $ValueToCheck, $txtCity)
	_VerifyData($DSN, $Acu_Filename, $STATE, $ConditionField, $ConditionOperator, $ValueToCheck, $txtState)
	_VerifyData($DSN, $Acu_Filename, $Zip, $ConditionField, $ConditionOperator, $ValueToCheck, $txtZip)

	; tab two data validation
	_VerifyData($DSN, $Acu_Filename, $ModelYear, $ConditionField, $ConditionOperator, $ValueToCheck, $txtModelYear)
	_VerifyData($DSN, $Acu_Filename, $PartsMarkup, $ConditionField, $ConditionOperator, $ValueToCheck, $txtPartsMarkup)
	_VerifyData($DSN, $Acu_Filename, $BCM_OXLO, $ConditionField, $ConditionOperator, $ValueToCheck, $txtBCM_OXLO)


	; Change the entry fields in preparation for Tab Change checking.
	_logger(@LF & "Changing Tab 1 fields to check data after tab." & @LF)
	_SetTab1_Fields()

	
	_logger("Changing Tabs")
	_ChangeTabs($TabControl, "R", $CurrentTab)
	_logger("Tab # after change = " & $CurrentTab)
	Sleep(1000)
	
	_logger(@LF & "Changing the Model Year to check data after tab. Value = 1901")
	_EnterTheText ("", "", $txtModelYear, "1901", 4 )
	_logger("Change Back and check PA field")
	_ChangeTabs($TabControl, "L", $CurrentTab)
	Sleep(1000)
	
	_Check_Tab1()
	
    if _ValueCompare($txtPACode, "12345") = "Equal" then  	
		_logger("Data remains through tabchange.")	   
;	$ReturnValue = _ValueCompare($txtPACode, "12345")
;	if $ReturnValue = "Equal" Then
;		_logger("Data remains through tabchange.")
	Else
		_logger("**** DATA NOT HOLDING THROUGH TABCHANGE on PA Code Data ****")
	EndIf
	 
	_logger("Change Back and check Model Year field")
	_ChangeTabs($TabControl, "R", $CurrentTab)
	Sleep(1000)
	$ReturnValue = _ValueCompare($txtModelYear, "1901")
	if $ReturnValue = "Equal" Then
		_logger("Data remains through tabchange.")
	Else
		_logger("**** DATA NOT HOLDING THROUGH TABCHANGE ****")
	EndIf 

		_logger("**** SAVE BUTTON TEST****")
   _CheckSaveButton($DSN, $Acu_FileName, $ConditionField, $ConditionOperator, $ValueToCheck )

	;_Acu_DropList("[CLASS:ComboBox; INSTANCE:1]", "")
	;_ComboBoxSelect("[CLASS:ComboBox; INSTANCE:1]", "")
 EndFunc
 
 Func _CheckSaveButton($DSN, $Acu_FileName, $ConditionField, $ConditionOperator, $ValueToCheck )
	_EnterTheText ("", "", $txtPACode, "12345", 4 )
	_EnterTheText ("", "", $txtPACheck, "&", 4 )		
	$SaveButton = "[CLASS:Button; INSTANCE:27]"
	_ClickTheButton ( "", "", $SaveButton, 4 )
	$strPA_Code = "FD_PA_CODE"
	$strPA_Check = "FD_CHECK_DIGIT"	
	
	$PA_COde = _GetData($DSN, $Acu_FileName, $strPA_Code, $ConditionField, $ConditionOperator, $ValueToCheck)

 EndFunc



Func _LoadArrays($Screen, $Field, $Value)

EndFunc

Func _VerifyData($DSN, $Acu_Filename, $FieldName, $ConditionField, $ConditionOperator, $ValueToCheck, $FieldInstance)
	$FileData = _GetData($DSN, $Acu_FileName, $FieldName, $ConditionField, $ConditionOperator, $ValueToCheck)
	$Result = _ValueCompare ($FieldInstance, $FileData)
	   $ControlText = ControlGetText("", "", $FieldInstance)
	IF $Result = "Equal" THEN 
		_logger($FieldName & " " & $Result)
		;_logger($FieldName & " " & $FileData & " " & $Result & " " & $ControlText)
	Else
		_logger( "*************** " & $FieldName & ' "' & $FileData & '" ' & $Result & ' "' & $ControlText & '"')
	EndIF	
	
EndFunc

Func _VerifyDistribution($DSN, $Acu_Filename, ByRef $FieldNames, $ConditionField, $ConditionOperator, $ValueToCheck, $FieldInstance)
;~ 	if IsArray($FieldNames) Then
;~ 		For
	$FileData = _GetData($DSN, $Acu_FileName, $FieldNames[0], $ConditionField, $ConditionOperator, $ValueToCheck)
	$FileData = $FileData & _GetData($DSN, $Acu_FileName, $FieldNames[1], $ConditionField, $ConditionOperator, $ValueToCheck)
	$FileData = $FileData & _GetData($DSN, $Acu_FileName, $FieldNames[2], $ConditionField, $ConditionOperator, $ValueToCheck)
	$Result = _ValueCompare ($FieldInstance, $FileData)
    $ControlText = ControlGetText("", "", $FieldInstance)
	IF $Result = "Equal" THEN 
	 ;_logger($FieldName & " " & $Result)
		_logger("Distribution Code" & " " & $FileData & " " & $Result & " " & $ControlText)
	Else
		_logger( "*************** Distribution Code" & ' "' & $FileData & '" ' & $Result & ' "' & $ControlText & '"')
	EndIF	
	
EndFunc
 
Func _SetTab1_Fields()
	
	If _EnterTheText ("", "", $txtPACode, "12345", 4 ) = 1 Then
		_logger("An error occurred changing PA Code")
	Else
		_logger("PA Code = 12345")
	EndIf
	
	If _EnterTheText('','',$txtPACheck, "%",4) = 1 Then
		_logger("An error occurred changing PA Check")
	Else	
		_logger('Check Digit = "%"' & @LF)
	EndIf
	
	If _EnterTheText('','',$txtSubDealer, "54321",4) = 1 Then
		_logger("An error occurred changing Sub Dealer")
	Else	
		_logger('Sub Dealer = "54321"' & @LF)
	EndIf
	
	If _EnterTheText('','',$txtSalesArea, 'BBB',4) = 1 Then
		_logger("An error occurred changing Geo Sales Area")
	Else	
		_logger('Geo Sales Area = "BBB"' & @LF)
	EndIf
	
	If _EnterTheText('','',$txtCurrency, 'CAN',4) = 1 Then
		_logger("An error occurred changing Currency")
	Else	
		_logger('Currency = "CAN"' & @LF)
	EndIf
	
	If _EnterTheText('','',$txtFordDistributionCode, "12E456", 4) = 1 Then
		_logger("An error occurred changing F-Dist")
	Else	
		_logger('F-Dist = "12E456"' & @LF)
	EndIf
	
	If _EnterTheText('','', $txtLMDistributionCode, "65E321",4) = 1 Then
		_logger("An error occurred changing LM-Dist")
	Else	
		_logger('LM-Dist = "65E321"' & @LF)
	EndIf
	
;~ 	Check this a different way in cahoots with the Division ID cbo  -->  $cboFranchiseCode
;~ 	$cboDivisionID

	If _EnterTheText('','', $txtCurrentLaborRate, "25.25",4) = 1 Then
		_logger("An error occurred changing Current Labor Rate")
	Else	
		_logger('Current Labor Rate = "25.25"' &@LF)
	EndIf
	
	If _EnterTheText('','', $txtPriorLaborRate, "75.75",4) = 1 Then
		_logger("An error occurred changing Prior Labor Rate")
	Else	
		_logger('Prior Labor Rate = "75.75"' &@LF)
	EndIf
	
;~ 	Check this a different way  -->  $dtEffectiveLaborDate
	
	If _EnterTheText('','', $txtDealerCode, "98765",4) = 1 Then
		_logger("An error occurred changing Dealer Code")
	Else	
		_logger('Dealer Code = "98765"' & @LF)
	EndIf
	
	If _EnterTheText('','', $txtDealerName, "Your Face Ford",4) = 1 Then
		_logger("An error occurred changing Name")
	Else	
		_logger('Dealer Name = "Your Face Ford"' & @LF)
	EndIf
	
	If _EnterTheText('','', $txtAddress, "Ford Place",4) = 1 Then
		_logger("An error occurred changing Address")
	Else	
		_logger('Dealer Addr = "Ford Place"' & @LF)
	EndIf
	
	If _EnterTheText('','', $txtCity, "Ford Town",4) = 1 Then
		_logger("An error occurred changing City")
	Else	
		_logger('Dealer City = "Ford Town"' & @LF)
	EndIf
	
	If _EnterTheText('','', $txtState, "Ford Country",4) = 1 Then
		_logger("An error occurred changing State")
	Else	
		_logger('Dealer Country = "Ford Country"' & @LF)
	EndIf
	
	If _EnterTheText('','', $txtZip, "Zip Ford",4) = 1 Then
		_logger("An error occurred changing Zip")
	Else	
		_logger('Dealer Zip = "Zip Ford"' & @LF)
	EndIf
	
EndFunc

Func _Check_Tab1()
	if not _ValueCompare($txtPACode, "12345") = "Equal" then  	
		_logger("**** PA Code Data Lost on TabChange ****")
	EndIf

	if not _ValueCompare($txtSubDealer, "54321") = "Equal" then  	
		_logger("**** SubDealer Data Lost on TabChange ****")
	EndIf
	
	if not _ValueCompare($txtSalesArea, "BBB") = "Equal" then  	
		_logger("**** SubDealer Data Lost on TabChange ****")
	EndIf

EndFunc 



