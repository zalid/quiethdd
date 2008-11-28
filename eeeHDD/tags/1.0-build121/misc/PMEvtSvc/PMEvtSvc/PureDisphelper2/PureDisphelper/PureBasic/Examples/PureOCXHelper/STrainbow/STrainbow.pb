; STrainbow freeware, download: http://www.st-software.at/activex.php

XIncludeFile "..\Register_Unregister_ActiveX.pb"

EnableExplicit

dhToggleExceptions(#True)

Define.l oRainbow1, oRainbow2, i

If OpenWindow(0, #PB_Ignore, #PB_Ignore, 300, 75, "Rainbow") And CreateGadgetList(WindowID(0))
 
  oRainbow1 = OCX_CreateGadget(1, 5, 5, 290, 30, "Project1.STprogressBar")
  oRainbow2 = OCX_CreateGadget(2, 5, 40, 290, 30, "Project1.STprogressBar")
  
  If oRainbow1 = 0
    RegisterActiveX("STrainbowBar.ocx")
    oRainbow2 = OCX_CreateGadget(1, 0, 0, 300, 300, "Project1.STprogressBar")
    oRainbow2 = OCX_CreateGadget(2, 0, 0, 300, 300, "Project1.STprogressBar")
  EndIf
  
  If oRainbow1
    dhPutValue(oRainbow1, "Styling = %d", 4)
    dhPutValue(oRainbow1, "UserLeftColor = %d", RGB(0,255,255))
    dhPutValue(oRainbow1, "UserRightColor = %d", RGB(255,0,0))
  EndIf

  If oRainbow2
    dhPutValue(oRainbow2, "Styling = %d", 4)
    dhPutValue(oRainbow2, "UserStyle = %d", 1)
    dhPutValue(oRainbow2, "UserLeftColor = %d", RGB(0,255,0))
    dhPutValue(oRainbow2, "UserRightColor = %d", RGB(255,0,0))
    dhPutValue(oRainbow2, "BackColor = %d", RGB(0,0,255))
    dhPutValue(oRainbow2, "PercentColor = %d", RGB(255,255,0))
  EndIf
  
  For I = 0 To 100
    dhPutValue(oRainbow1, "Value = %d", I)
    dhPutValue(oRainbow2, "Value = %d", 100 - I)
    
    While WindowEvent() : Wend
    Delay(200)
    
  Next
   
  While WaitWindowEvent() ! 16 : Wend
 
  CloseWindow(0)

  If oRainbow1 : dhReleaseObject(oRainbow1) : EndIf
  If oRainbow2 : dhReleaseObject(oRainbow2) : EndIf
EndIf

; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 22
; FirstLine = 6
; Folding = -
; EnableXP
; HideErrorLog