
Define DISPHELPER_NO_FOR_EACH
Define DISPHELPER_NO_OCX_CreateGadget

XIncludeFile "DispHelper_Include.pb"

dhInitializeImp()

Define.l cHeight, cWidth

If ExamineDesktops()
  cWidth  = DesktopWidth(0)   / 2 - 50
  cHeight = DesktopHeight(0)  / 4
EndIf

dhToggleExceptions(#True)

Define.l oAgent, oGenie, oChar

oAgent = dhCreateObject("Agent.Control.1")

If oAgent
  dhPutValue(oAgent, "Connected = %b", #True)
  dhCallMethod(oAgent, "Characters.Load(%T)", @"Genie")
  dhGetValue("%o", @oGenie, oAgent, "Characters(%T)", @"Genie")
  If oGenie
    dhCallMethod(oGenie, "Show")
    Delay(3000)
    dhCallMethod(oGenie, "MoveTo(%d,%d)", cWidth, cHeight)
    dhCallMethod(oGenie, "Play(%T)", @"Greet")
    dhCallMethod(oGenie, "Speak(%T)", @"Hello, feel the ..Pure.. Power of PureBasic")
    dhCallMethod(oGenie, "Play(%T)", @"Reading")
    Delay(20000)
    dhCallMethod(oGenie, "Stop")
    dhCallMethod(oGenie, "Speak(%T)", @"PureBasic is a nice computer language")
    Delay(6000)
    MessageRequester("Agent", "click ok to end")
    dhCallMethod(oGenie, "Play(%T)", @"Hide")
    Delay(3000)
    dhReleaseObject(oGenie)
  EndIf
  dhReleaseObject(oAgent)
EndIf

dhUninitialize()
; IDE Options = PureBasic 4.20 Beta 2 (Windows - x86)
; CursorPosition = 24
; Folding = -
; EnableXP
; HideErrorLog