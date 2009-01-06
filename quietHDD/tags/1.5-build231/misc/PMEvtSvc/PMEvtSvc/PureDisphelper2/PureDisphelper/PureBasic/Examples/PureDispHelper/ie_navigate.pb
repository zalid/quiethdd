; based on FreeBASIC example
; changed to PB by ts-soft

EnableExplicit

Define.l ieApp

dhToggleExceptions(#True)

ieApp = dhCreateObject("InternetExplorer.Application")

If ieApp

  dhPutValue (ieApp, "Visible = %b", #True)
  dhCallMethod(ieApp, "Navigate(%T)", @"www.purebasic.com")
  dhReleaseObject(ieApp) : ieApp = 0

EndIf

; IDE Options = PureBasic 4.20 Beta 2 (Windows - x86)
; Folding = -
; EnableXP
; EnableUser
; HideErrorLog