EnableExplicit

dhToggleExceptions(#True)

Define.l oQT
Define.s MediaPath = "http://www.blitzbasement.net/bank/aliensong.mpeg"

If OpenWindow(0, #PB_Ignore, #PB_Ignore, 320, 255,"QuickTime - Demo")

  CreateGadgetList(WindowID(0))
  oQT = OCX_CreateGadget(0, 0, 0, 320, 255, "QTOControl.QTControl.1")

  If oQT

    dhPutValue(oQT, "FileName = %T", @MediaPath)
    dhPutValue(oQT, "AutoPlay = %b", #True)

  EndIf

  While WaitWindowEvent() ! 16 : Wend

  If oQT : dhCallMethod(oQT, "ShowAboutBox") : dhReleaseObject(oQT) : EndIf

  CloseWindow(0)
EndIf
