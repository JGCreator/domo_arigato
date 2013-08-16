
Func _OpenPartsInvoice()
   SEND("{ALT}")
   Sleep(1000)
   SEND("{P}")
   Sleep(1000)
   SEND("{T}")
   Sleep(1000)
   SEND("{ENTER}")
EndFunc

Func _OpenPriceCodes()
   SEND("{ALT}")
   Sleep(1000)
   SEND("{P}")
   Sleep(1000)
   SEND("{C}")
EndFunc   

Func _OpenCustMaint()
   SEND("{ALT}")
   Sleep(1000)
   SEND("{M}")
   Sleep(1000)
   SEND("{C}")   
EndFunc

Func _OpenPartsMaint()
   SEND("{ALT}")
   Sleep(1000)
   SEND("{P}")
   Sleep(1000)
   SEND("{T}")
   Sleep(1000)
   SEND("{T}")
   Sleep(1000)
   SEND("{ENTER}")
EndFunc

Func _OpenFordDealerInfo()
	ControlFocus("", "", "[CLASS:Edit; INSTANCE:1]")
		send("frmFordDealerInfo")
		send("{ENTER}")
		
EndFunc
	