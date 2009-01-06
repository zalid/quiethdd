XIncludeFile "Register_Unregister_ActiveX.pb"

EnableExplicit

Define.l oAU3, Result

oAU3 = dhCreateObject("AutoItX3.Control")

If oAU3 = #False
  If RegisterActiveX("AutoItX3.dll")
    oAU3 = dhCreateObject("AutoItX3.Control")
  EndIf
EndIf

If oAU3
  MessageRequester("AutoItX3", "This script will run some test calculations")
  dhToggleExceptions(#True)
  dhGetValue("%d", @Result, oAU3, "Run(%T,%T,%d",@"calc.exe", @"", #SW_SHOWNORMAL)
  dhGetValue("%b", @Result, oAU3, "WinWaitActive(%T,%T,%d",@"Calculator", @"", 5)
  If Result
    dhGetValue("%b", @Result, oAU3, "Send(%T)",@"2*2=")
    dhCallMethod(oAU3, "Sleep = %d", 500)
    dhGetValue("%b", @Result, oAU3, "Send(%T)",@"4*4=")
    dhCallMethod(oAU3, "Sleep = %d", 500)
    dhGetValue("%b", @Result, oAU3, "Send(%T)",@"8*8=")
    dhCallMethod(oAU3, "Sleep = %d", 1000)
    dhGetValue("%b", @Result, oAU3, "WinClose(%T,%T)", @"Calculator", @"")
  EndIf
  dhReleaseObject(oAU3)
EndIf
; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 26
; Folding = -
; EnableXP
; HideErrorLog