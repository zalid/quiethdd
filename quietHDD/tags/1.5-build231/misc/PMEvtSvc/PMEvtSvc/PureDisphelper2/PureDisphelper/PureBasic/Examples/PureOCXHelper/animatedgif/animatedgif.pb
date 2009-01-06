XIncludeFile "..\Register_Unregister_ActiveX.pb"

; http://www.freevbcode.com/ShowCode.asp?id=4045

; example by Kiffi

EnableExplicit

Enumeration
  #Load
  #Unload
  #Stop
  #Restart
  #NextFrame
  #GIF
EndEnumeration

Define.l oGIF

dhToggleExceptions(#True)

If OpenWindow(0, #PB_Ignore, #PB_Ignore, 300, 300, "AniGIF-Demo")
  If CreateGadgetList(WindowID(0))
    
    ButtonGadget(#Load, 5, 5, 70, 25, "Load...")
    ButtonGadget(#Unload, 5, 35, 70, 25, "Unload")
    ButtonGadget(#Stop, 5, 65, 70, 25, "Stop")
    ButtonGadget(#Restart, 5, 125, 70, 25, "Restart")
    ButtonGadget(#NextFrame, 5, 95, 70, 25, "Next frame")
    
    oGIF = OCX_CreateGadget(#GIF, 80, 5, 215, 290, "AnimatedGif.AniGif")
    
    If oGIF = 0
      RegisterActiveX("AnimatedGif.ocx")
      oGIF = OCX_CreateGadget(#GIF, 80, 5, 215, 290, "AnimatedGif.AniGif")
    EndIf
    
    If oGIF = 0 
      MessageRequester("", "Please register animatedgif.ocx before using")
      End
    EndIf
    
    dhCallMethod(oGIF, ".LoadFile(%T, %b)", @"computer16.gif", #False)
    
    Define.s myFile
    
    Repeat
      
      Select WaitWindowEvent()
        
        Case #PB_Event_Gadget
          Select EventGadget()
            
            Case #Load
              myFile=OpenFileRequester("Please select", myFile, "GIF files (*.gif)|*.gif", 1)
              If myFile
                dhCallMethod(oGIF, ".LoadFile(%T, %b)", @myFile, #False)
              EndIf
              
            Case #Unload
              dhCallMethod(oGIF, ".LoadFile(%T, %b)", @"", #False)
              
            Case #Stop
              dhCallMethod(oGIF, ".StopAnimate(%b, %b)", #True, #True)
              
            Case #Restart
              dhCallMethod(oGIF, ".RestartAnimate()")
              
            Case #NextFrame
              dhCallMethod(oGIF, ".NextFrame()")
              
          EndSelect
          
        Case #PB_Event_CloseWindow
          Break
          
      EndSelect
      
    ForEver
    
    CloseWindow(0)
    
    If oGIF : dhReleaseObject(oGIF) : EndIf
    
  EndIf
EndIf

; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 60
; Folding = -
; HideErrorLog