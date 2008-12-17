; STclock freeware, download: http://www.st-software.at/activex.php

XIncludeFile "..\Register_Unregister_ActiveX.pb"

dhToggleExceptions(#True)

Define.l oClock

If OpenWindow(0, #PB_Ignore, #PB_Ignore, 300, 300, "Clock") And CreateGadgetList(WindowID(0))
 
  oClock = OCX_CreateGadget(1, 0, 0, 300, 300, "projClock.STclock")
  
  If oClock = 0
    RegisterActiveX("STclock.ocx")
    oClock = OCX_CreateGadget(1, 0, 0, 300, 300, "projClock.STclock")
  EndIf
  
  If oClock
    dhPutValue(oClock, "ColorHour   = %d", $FF0000)
    dhPutValue(oClock, "Colorminute = %d", $000000)
    dhPutValue(oClock, "Colorsecond = %d", $0000FF)
    dhPutValue(oClock, "ForeColor   = %d", $3B642C)
  EndIf

  While WaitWindowEvent() ! 16 : Wend
  
  dhCallMethod(oClock, "ShowAbout")
  
  CloseWindow(0)

  If oClock : dhReleaseObject(oClock) : EndIf
  
EndIf

; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 1
; Folding = -
; EnableXP
; HideErrorLog