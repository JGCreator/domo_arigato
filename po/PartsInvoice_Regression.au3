#include "//ddsdevsc/u/Automated_Testing/Buttonclick.au3"
#include "//ddsdevsc/u/Automated_Testing/OpenProgramHotKeys.au3"
Local $file = FileOpen("test.txt", 2)
If $file = -1 Then
    MsgBox(0, "Error", "Unable to open file.")
    Exit
EndIf
AutoItSetOption("WinTitleMatchMode", 4)	 
AutoItSetOption("SendKeyDelay", 75)
Global $Vendor = InputBox("", "Please enter vendor and click OK")
If @Error = 1 Then ; cancel pushed
	Exit
EndIf
;$Vendor = "98"
;$PriceCode = "CRM"
;$PriceCodeDesc = "TEST"
;$FirstFactor = "F-1000"
;$SecondFactor = "L-2000"
;$Part = "90341C3"
;$PartListPrice = "100.00"
;$PartFleetPrice = "200.00"
;$PartQuantity = "1"
;$Customer = "47"
;$Initials = "DDS"
Global $PriceCode = InputBox("", "Please enter 3 chars for price code and click OK")
If @Error = 1 Then ; cancel pushed
	Exit
EndIf

Global $PriceCodeDesc = InputBox("", "Please enter your price code description and click OK")
If @Error = 1 Then ; cancel pushed
	Exit
EndIf

Global $FirstFactor = InputBox("", "Please enter first factorand click OK" & @CRLF & "(ex. Fleet - 10% = F-1000)")
If @Error = 1 Then ; cancel pushed
	Exit
EndIf

Global $SecondFactor = InputBox("", "Please enter second factor and click OK")
If @Error = 1 Then ; cancel pushed
	Exit
EndIf

Global $Part = InputBox("", "Please enter part number and click OK")
If @Error = 1 Then ; cancel pushed
	Exit
EndIf

Global $PartListPrice = InputBox("", "Please enter part list price and click OK")
If @Error = 1 Then ; cancel pushed
	Exit
EndIf

Global $PartFleetPrice = InputBox("", "Please enter part fleet price and click OK")
If @Error = 1 Then ; cancel pushed
	Exit
EndIf

Global $PartQuantity = InputBox("", "Please enter part quantity and click OK")
If @Error = 1 Then ; cancel pushed
	Exit
EndIf

Global $Customer = InputBox("", "Please enter customer number and click OK")
If @Error = 1 Then ; cancel pushed
	Exit
EndIf

Global $Initials = InputBox("", "Please enter valid initials and click OK")
If @Error = 1 Then ; cancel pushed
	Exit
EndIf
_Main()

Func _Main()
   _WaitActivate("Advantage - Dubuque Data", 4)
   _OpenPriceCodes()
   _PriceCode()
   _WaitActivate("Advantage - Dubuque Data", 4)  
   _OpenCustMaint()
   _CustMaint()
   _WaitActivate("Advantage - Dubuque Data", 4)
   _OpenPartsMaint()
   _Parts_Maintenance()
   _WaitActivate("Advantage - Dubuque Data", 4)  
   _OpenPartsInvoice()
   _PartsInvoice()
   _PartsInvoice3()
   _WaitActivate("Advantage - Dubuque Data", 4)
   _OpenPartsMaint()
   _PartsMaint2()
   _WaitActivate("Advantage - Dubuque Data", 4)
   _OpenPartsInvoice()
   _PartsInvoice()
   _PartsInvoice2()
   Exit
EndFunc


#cs ----------------------------------------------------------------------------
   Creates a price code.
#ce-----------------------------------------------------------------------------        
Func _PriceCode()
   _WaitActivate("Price Code Maintenance", 1)
   _ClickTheButton("Price Code Maintenance", "", "[CLASS:Button; INSTANCE:1]", 1)
   _ClickTheButton("Price Code Maintenance", "", "[CLASS:Edit; INSTANCE:1]", 1)
   _EnterTheText("Price Code Maintenance", "", "[CLASS:Edit; INSTANCE:1]", $PriceCode, 1)
   ;_Send("{TAB}")
   _ClickTheButton("", "", "[CLASS:Button; INSTANCE:15]", 1)
   _WaitActivate("Price Codes", 1)
   Sleep(500)
   _SendTab($Vendor, 1)
   _SendTab("{TAB}", 1)
   _SendTab($PriceCodeDesc, 1)
   _SendTab("{TAB}", 1)
   _SendTab($FirstFactor, 1)
   _SendTab("{TAB}", 1)
   _SendTab($SecondFactor, 1)
   _SendTab("{TAB}", 1)
   _SendTab("N", 1)
   _ClickTheButton("", "", "[CLASS:Button; INSTANCE:2]", 1)
   Local $sText = WinGetTitle("[CLASS:AcuBodyClass; INSTANCE:1]", "")
   if $sText = "Price code already exists for vendor." then
        _ClickTheButton("", "", "[CLASS:Button; INSTANCE:1]", 1)
		_WaitActivate("Price Code Maintenance", 1)
        _ClickTheButton("", "", "[CLASS:Button; INSTANCE:2]", 1)
   else
        _ClickTheButton("", "", "[CLASS:Button; INSTANCE:1]", 1)
   Endif
   _WaitActivate("Price Code Maintenance", 1)
   _ClickTheButton("Price Code Maintenance", "", "[CLASS:Button; INSTANCE:14]", 1)
EndFunc


#cs ----------------------------------------------------------------------------
   Sets the Part List Price/Fleet Price and price factor.
#ce-----------------------------------------------------------------------------      
Func _Parts_Maintenance()
   _WaitActivate("Parts Master", 4)
   _SendTab($Part, 1)
   _SendTab("{TAB}", 1)
   _ClickTheButton("Parts Master", "", "[CLASS:SysTabControl32; INSTANCE:1]", 1)
   _ClickTheButton("Parts Master", "", "[CLASS:Edit; INSTANCE:12]", 1)
   _EnterTheText("Parts Master", "", "[CLASS:Edit; INSTANCE:12]", $PartListPrice, 1)
   _ClickTheButton("Parts Master", "", "[CLASS:Edit; INSTANCE:15]", 1)
   _EnterTheText("Parts Master", "", "[CLASS:Edit; INSTANCE:15]", $PartFleetPrice, 1)
   _ClickTheButton("Parts Master", "", "[CLASS:SysTabControl32; INSTANCE:1]", 2)
   _SendTab("{RIGHT}", 2)
   _ClickTheButton("Parts Master", "", "[CLASS:Edit; INSTANCE:50]", 1)
   _EnterTheText("Parts Master", "", "[CLASS:Edit; INSTANCE:50]", $PriceCodeDesc, 1)
   _ClickTheButton("Parts Master", "", "[CLASS:Button; INSTANCE:23]", 1)
EndFunc


#cs ----------------------------------------------------------------------------
   Enters initials & customer number. Opens Customer Maintenance
#ce-----------------------------------------------------------------------------      
Func _PartsInvoice()
   _WaitActivate("Parts Invoice", 4)
   _SendTab($Initials, 1)
   _SendTab("{TAB}", 1)
   _ClickTheButton("Parts Invoice", "", "[CLASS:Edit; INSTANCE:4]", 1)
   _SendTab($Customer, 1)
   _SendTab("{TAB}", 1)
EndFunc


#cs ----------------------------------------------------------------------------
   Modifies the customer's price code.
#ce-----------------------------------------------------------------------------     
Func _CustMaint()
   _WaitActivate("Customer Maintenance", 4)
   _SendTab($Customer, 2)
   _SendTab("{TAB}", 6)
   _ClickTheButton("Customer Maintenance", "", "[CLASS:Edit; INSTANCE:29]", 4)
   _EnterTheText("Customer Maintenance", "", "[CLASS:Edit; INSTANCE:29]", $PriceCode, 4)
   _SendTab("{TAB}", 2)
   Send("!s")
   Sleep(5000)
   Send("!x")
EndFunc

Func _PartsMaint2()
   _WaitActivate("Parts Master", 4)
   _SendTab($Part, 1)
   _SendTab("{TAB}", 1)
   _ClickTheButton("Parts Master", "", "[CLASS:SysTabControl32; INSTANCE:1]", 1)
   _ClickTheButton("Parts Master", "", "[CLASS:Edit; INSTANCE:15]", 1)
   _EnterTheText("Parts Master", "", "[CLASS:Edit; INSTANCE:15]", ".00", 1)
   _ClickTheButton("Parts Master", "", "[CLASS:Button; INSTANCE:23]", 1)
EndFunc

Func _PartsInvoice2()
   _ClickTheButton("Parts Invoice", "", "[CLASS:Edit; INSTANCE:7]", 1)
   _SendTab($Part, 1)
   _SendTab("{TAB}", 1)
   _SendTab($PartQuantity, 1)
   _SendTab("{TAB}", 4)
   Local $sText = ControlGetText("Parts Invoice", "", "[CLASS:Edit; INSTANCE:13]")
   if $sText = 80.00 then
        FileWriteLine($file,"Second Price Code Passed")
   else
        FileWriteLine($file, "Second Price Code Failed")
   Endif      
   _ClickTheButton("Parts Invoice", "", "[CLASS:Button; INSTANCE:26]", 1)
   _ClickTheButton("Parts Invoice", "", "[CLASS:Button; INSTANCE:1]", 1)
EndFunc 

Func _PartsInvoice3()
   _ClickTheButton("", "", "[CLASS:Edit; INSTANCE:7]", 1)
   _SendTab($Part, 1)
   _SendTab("{TAB}", 1)
   _SendTab($PartQuantity, 1)
   _SendTab("{TAB}", 4)
   Local $sText = ControlGetText("Parts Invoice", "", "[CLASS:Edit; INSTANCE:13]")
   if $sText = 180.00 then
        FileWriteLine($file,"First Price Code Passed")
   else
        FileWriteLine($file, "First Price Code Failed")
   Endif      
   _ClickTheButton("Parts Invoice", "", "[CLASS:Button; INSTANCE:26]", 1)
   _ClickTheButton("Parts Invoice", "", "[CLASS:Button; INSTANCE:1]", 1)
EndFunc

