; example by ts-soft

Define DISPHELPER_NO_FOR_EACH
Define DISPHELPER_NO_OCX_CreateGadget

XIncludeFile "DispHelper_Include.pb"

EnableExplicit

dhInitializeImp()

Define.l obj, Result

Define.s MediaPath = "http://www.blitzbasement.net/bank/aliensong.mpeg"

dhToggleExceptions(#True)

If OpenWindow(0, #PB_Ignore, #PB_Ignore, 600, 470, "MediaPlayer")

  obj = dhCreateObject("MediaPlayer.MediaPlayer", WindowID(0))
  
  If obj

    dhPutValue(obj, "FileName = %T", @MediaPath)
    dhPutValue(obj, "FileName = %T", @MediaPath)

  EndIf
  
  While WaitWindowEvent() <> #PB_Event_CloseWindow : Wend
  
  dhCallMethod(obj, "AboutBox")
  dhReleaseObject(obj)
  CloseWindow(0)
  
EndIf
; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 23
; Folding = -
; EnableUnicode
; HideErrorLog