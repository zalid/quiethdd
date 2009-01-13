; example by ts-soft

EnableExplicit

Define.l oRMChart
Define.s Pyramide = PeekS(?Pyramide)

dhToggleExceptions(#True)

If OpenWindow(0, #PB_Ignore, #PB_Ignore, 600, 450, "RMChart loaded from Memory") And CreateGadgetList(WindowID(0))
 
  oRMChart = OCX_CreateGadget(1, 0, 0, 600, 450, "RMChart.RMChartX")
 
  If oRMChart
    dhPutValue(oRMChart, ".RMCFile = %T", @Pyramide)
    dhCallMethod(oRMChart, ".Draw")
   
  EndIf
 
  While WaitWindowEvent() ! 16 : Wend
 
  CloseWindow(0)
 
  If oRMChart : dhReleaseObject(oRMChart) : EndIf
 
EndIf

DataSection
  Pyramide:
  IncludeBinary "Pyramide.rmc"
EndDataSection
; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 14
; Folding = -
; EnableXP
; HideErrorLog