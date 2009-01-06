; by ts-soft

EnableExplicit

dhToggleExceptions(#True)

Define.l obj = dhCreateObject("SAPI.SpVoice")

If obj

  dhCallMethod(obj, "Speak (%T,%d)", @"PureBasic, feel the pure Power", 0)
  dhReleaseObject(obj) : obj = 0

EndIf

; IDE Options = PureBasic 4.20 Beta 2 (Windows - x86)
; Folding = -
; EnableXP
; EnableUser
; HideErrorLog