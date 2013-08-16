; include this script only once in any series of scripts.
#include-once

;~ ; include the Common Control functions.
#include "./Includes/CommonControlFunctions.au3"


; Vars used to run as stand alone function.
;~  $Win_Text 		= "2013081309053194"
;~  $Sales_Code	= ''
;~  $Part_Nbr		= "SP417"
;~  $Qty			= 4
;~  _PartsInvoice_AddPart($Win_Text, $Sales_Code, $Part_Nbr, $Qty)


#cs----
Description: 
	Add a part to the Parts Invoice using a provided quantity.
	
	_PartsInvoice_AddPart("Win Text", "Part #"[, "Qty"[, "Sales Code"[, "Popup option"]]])
	
params:
	$Win_Text:		The time stamp found in the label of parts invoice (low right corner / invisible)
	$Part_Nbr:		Part number to be added
	$Qty:			Quantity used for part added ( default 1 if none supplied )
	$Sales_Code:	(optional)		Specific sales code to be used 
	$popup:			(optional)		Action to take on popup boxes ( defalut is 2 )
	  +---->	(add the values together)
						1 	- Add to Parts master & Sell core
						2 	- Add to Parts master & Exchange core
						3   - Don't add to Parts master & Sell core
						4   - Don't add to Parts master & Exchange core
		
		$Vendor:	The vendor number the part should be added for		
		$Status:	Parts master status of the part
		$Bin:		Bin location of the part
	
Returns:
	0 	= success
	1  	= fail - generic
	91	= fail - too many parts
	-1	= fail - nonmaster part not handled
	-2 	= fail - core part not handled		

#ce----
Func _PartsInvoice_AddPart(	$Win_Text, _ 
							$Part_Nbr, _ 
							$Qty = 1, _
							$Sales_Code = ' ', _
							$PopUp = 0, _ 
							$addBin = "<SYS>", _
							$addVendor = '', _
							$addStatus = "T" )

; Allow partial match of window title
Opt("WinTitleMatchMode", 4)
opt("WinTextMatchMode", 1)

	
; verification / default of function arguments.
	; convert to string in case not.
	$Win_Title 	= "Parts Invoice" ;string($Win_Title)
	$Win_Text 	= string($Win_Text)
	
	; assume qty of 1 if not number
	If Not IsNumber($Qty) Then
		$Qty = 1
	EndIf
	
	; sleep until window exists
	While Not WinExists ( $Win_Title, $Win_Text )
		Sleep ( 1000 )
	WEnd
	
	; activate the window.
	If Not WinActive ( $Win_Title, $Win_Text ) Then WinActivate ( $Win_Title, $Win_Text )

	; define and initialize values for window controls
	$txtSale	= "[CLASS:Edit; INSTANCE:6]"	; sales code entry
	$txtPart 	= "[CLASS:Edit; INSTANCE:7]"	; part number entry
	$txtQty 	= "[CLASS:Edit; INSTANCE:8]"	; quantity entry
	$btnSave	= "[CLASS:Button; INSTANCE:17]"	; save&new button
	
	; if sales code param exists put it
	If $Sales_Code <> '' And $Sales_Code <> " " Then
		ConsoleWrite("Sales Code IS NOT null or space, attempting to put text" & @LF & "Sales Code = " & $Sales_Code & @LF)
		; try calling enter text function
		If _EnterTheText ( $Win_Title, $Win_Text, $txtSale, $Sales_Code, .65 ) = 1 Then
			; bail on error / time-out
			ConsoleWrite("Put text failed, return 1" & @LF)
			Return 1			
		EndIf
		; tab after success
		ConsoleWrite("tabbing to next" & @LF)
		Send("{TAB}")
	Else
		; tab if no param value
		ConsoleWrite("Sales Code IS null or space, tabbing to next" & @LF)
		Send("{TAB}")
	EndIf
	
	
	; enter the part number text
	If _EnterTheText ( $Win_Title, $Win_Text, $txtPart, $Part_Nbr, .65 ) = 1 Then
		; bail on error / time-out
		ConsoleWrite("Put Part number failed" & @LF)
		Return 1
	Else
		; tab after success
		ConsoleWrite("Part number entered successfully, tabbing to next" & @LF)
		Send("{TAB}")
	EndIf
	
	
	; check for popups
	sleep( 150 )	
	If WinExists("Part Not Found", "No Part Master Record") _ 
		And $PopUp > 0 Then
		ConsoleWrite('non-mast popup found')
		WinActivate("Part Not Found", "No Part Master Record")
		_PopUp_Box($PopUp, "Non-Master")
	Else
		; return fail code for no master
		Return -1
	EndIf
	
;~ 	If WinExists("
	
	
	; enter the quantity text
	If _EnterTheText ( $Win_Title, $Win_Text, $txtQty, $Qty, .65 ) = 1 Then
		ConsoleWrite("Put Quantity failed" & @LF)
		Return 1
	Else
		ConsoleWrite("Quantity entered successfully, tabbing to next" & @LF)
		Send("{TAB}")
	EndIf
	
	ConsoleWrite("Clicking the button" & @LF)
	_ClickTheButton ( $Win_Title, $Win_Text, $btnSave, .5 )
	
	
	
	; check for a popup box.
	sleep( 150 )
	If WinExists("Core Part Info") And $PopUp > 0 Then
		WinActivate("Core Part Info")
		_PopUp_Box($PopUp, "Core-Part")
	Else
		Return -2
	EndIf
	ConsoleWrite("END")
EndFunc

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
Func _PopUp_Box($Action, $Type)
	
	$btn_Yes_Sell	= "[CLASS:Button; INSTANCE:1]"
	$btnNo_Exchange = "[CLASS:Button; INSTANCE:2]"
	
	Select 
	Case $Action = 1
			_ClickTheButton ( '', '', $btn_Yes_Sell, .25 )
			Return 0
	Case $Action = 2
			If $Type = "Non-Master" Then
				_ClickTheButton ( '', '', $btn_Yes_Sell, .25 )
				Return 0
			Else
				_ClickTheButton ( '', '', $btnNo_Exchange, .25 )
				Return 0
			EndIf
	Case $Action = 3
			If $Type = "Non-Master" Then
				_ClickTheButton ( '', '', $btnNo_Exchange, .25 )
				Return 0
			Else
				_ClickTheButton ( '', '', $btn_Yes_Sell, .25 )
				Return 0
			EndIf
	Case $Action = 4
			_ClickTheButton ( '', '', $btnNo_Exchange, .25 )
			Return 0
			
	Case Else
		Return 1
		
	EndSelect
	
EndFunc

