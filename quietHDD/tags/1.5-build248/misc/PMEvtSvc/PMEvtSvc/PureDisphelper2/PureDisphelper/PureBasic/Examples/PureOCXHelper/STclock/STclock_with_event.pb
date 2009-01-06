; STclock freeware, download: http://www.st-software.at/activex.php

XIncludeFile "..\Register_Unregister_ActiveX.pb"
XIncludeFile "..\ComEventSink.pb"

dhToggleExceptions(#True)

Global oClock.l

Procedure EventCallback(Event$, ParameterCount, *Params)
; Event Click
; Occurs when the user presses And then releases a mouse button over an object.
; 
; Event DblClick
; Occurs when the user presses And releases a mouse button And then presses And releases it again over an object.
; 
; Event KeyDown(KeyCode As Integer, Shift As Integer)
; Occurs when the user presses a key While an object has the focus.
; 
; Event KeyPress(KeyAscii As Integer)
; Occurs when the user presses And releases an ANSI key.
; 
; Event KeyUp(KeyCode As Integer, Shift As Integer)
; Occurs when the user releases a key While an object has the focus.
; 
; Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
; Occurs when the user presses the mouse button While an object has the focus.
; 
; Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
; Occurs when the user moves the mouse.
; 
; Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
; Occurs when the user releases the mouse button While an object has the focus.

  Select Event$
    Case "Click"
      Debug "Click"
    Case "DblClick"
      Debug  "Double Click"
      dhCallMethod(oClock, "ShowAbout")
    Case "KeyPress"
      Debug "You pressed: " + Chr(OCX_EventLong(*Params, 1))
  EndSelect
  
EndProcedure


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
    
    OCX_ConnectEvents(oClock, @EventCallback())
  EndIf

  While WaitWindowEvent() ! 16 : Wend

  CloseWindow(0)

  If oClock : dhReleaseObject(oClock) : EndIf
  
EndIf

; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 37
; FirstLine = 26
; Folding = -
; EnableXP
; HideErrorLog