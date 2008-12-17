; example by ts-soft

EnableExplicit

XIncludeFile "DispHelper_Include.pb"

Procedure.s GetExePath()
  Protected Result.s

  If UCase(GetPathPart(ProgramFilename())) = UCase(#PB_Compiler_Home + "Compilers\")
    Result = GetCurrentDirectory()
  Else
    Result = GetPathPart(ProgramFilename())
  EndIf

  If Right(Result, 1) <> "\" : Result + "\" : EndIf

  ProcedureReturn Result
EndProcedure

dhInitializeImp()

Define.l oFlash, Result
Define.s Movie = GetExePath() + "worm.swf"

If Movie
  If OpenWindow(0, #PB_Ignore, #PB_Ignore, 440, 280, "Flash-Demo") And CreateGadgetList(WindowID(0))
  
    ContainerGadget(0, 0, 0, 440, 240)
    CloseGadgetList()
    ButtonGadget(1, 20, 250, 60, 25, "Run")
    ButtonGadget(2, 90, 250, 60, 25, "Stop")
    
    dhToggleExceptions(#True)
    oFlash = dhCreateObject("ShockwaveFlash.ShockwaveFlash", GadgetID(0))
    
    If oFlash
      
      dhCallMethod(oFlash, "LoadMovie (%b,%s)", #False, @Movie)
      
      Repeat
        
        dhGetValue("%d", @Result, oFlash, "ReadyState")
 
      Until Result = 4
       
      dhCallMethod(oFlash, "Play")
           
      Repeat
        
        Select WaitWindowEvent()
        
          Case #PB_Event_CloseWindow
            Break
          
          Case #PB_Event_Gadget
          
            Select EventGadget()
            
              Case 1
                dhCallMethod(oFlash, "Play")
                
              Case 2
                dhCallMethod(oFlash, "Stop")
                
            EndSelect
            
        EndSelect
      
      ForEver
      
      dhReleaseObject(oFlash)
    
    EndIf
    
    CloseWindow(0)
    
  EndIf
  
EndIf

dhUninitialize()

; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 82
; FirstLine = 35
; Folding = -
; EnableXP
; HideErrorLog