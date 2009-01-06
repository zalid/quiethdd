
EnableExplicit

Define oWScript = dhCreateObject("WScript.Shell")

dhToggleExceptions(#True)

If oWScript
  dhCallMethod(oWScript, "Run(%T)", @"calc")
  Delay(100)
  dhCallMethod(oWScript, "AppActivate(%T)", @"Calculator"); english
  ;dhCallMethod(oWScript, "AppActivate(%T)", @"Rechner"); german
  Delay(100)
  dhCallMethod(oWScript, "SendKeys(%T)", @"1{+}")
  Delay(500)
  dhCallMethod(oWScript, "SendKeys(%T)", @"2")
  Delay(500)
  dhCallMethod(oWScript, "SendKeys(%T)", @"~")
  Delay(500)
  dhCallMethod(oWScript, "SendKeys(%T)", @"*3")
  Delay(500)
  dhCallMethod(oWScript, "SendKeys(%T)", @"~")
  Delay(2500)
  dhCallMethod(oWScript, "SendKeys(%T)", @"%{F4}")
  dhReleaseObject(oWScript)
EndIf

; IDE Options = PureBasic 4.20 Beta 2 (Windows - x86)
; Folding = -
; EnableXP
; EnableUser
; HideErrorLog