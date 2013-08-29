#include-once
#include "CommonControlFunctions.au3"
#cs
Author:  TGUZ & JGUST
Date: 08/21/13
   Description:  This function is login to Advantage, the variables User Name, User Password and Site need to be declared inside of your script.
#ce

opt('WinTitleMatchMode', 1)
Func _Advantage_Login($UserName, $UserPassword, $Site)
	WinWait("Advantage Login", '', 2)
	if not WinActive ("Advantage Login") Then WinActivate ("Advantage Login")
	_EnterTheText("Advantage Login","","[CLASS:Edit; INSTANCE:1]",$UserName,.5)
	_EnterTheText("Advantage Login","","[CLASS:Edit; INSTANCE:2]",$UserPassword,.5)
	_EnterTheText("Advantage Login","","[CLASS:Edit; INSTANCE:3]",$Site,.5)
	Send("{TAB}")
	_ClickTheButton("Advantage Login","OK","[CLASS:Button; INSTANCE:1]",1)
EndFunc