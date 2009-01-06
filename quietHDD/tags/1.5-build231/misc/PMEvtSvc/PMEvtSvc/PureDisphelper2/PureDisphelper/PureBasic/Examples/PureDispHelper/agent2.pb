; example by ts-soft

EnableExplicit

dhToggleExceptions(#True)

Define.l oAgent, oMerlin
Define.s item

oAgent = dhCreateObject("Agent.Control.2")

If oAgent = 0
  Debug "error: can't create object"
  End
EndIf

dhPutValue(oAgent, "Connected = %b", #True)
dhCallMethod(oAgent, "Characters.Load(%T,%T)", @"Merlin", @"Merlin.acs")
dhGetValue("%o", @oMerlin, oAgent, "Characters(%T)", @"Merlin")

If oMerlin = 0
  Debug "error: can't create object"
  End
EndIf

dhCallMethod(oMerlin, "Show")

If OpenWindow(0, 0, 0, 200, 200, "Merlin - Demo", #PB_Window_ScreenCentered|#PB_Window_SystemMenu) And CreateGadgetList(WindowID(0))

  ListViewGadget(0, 5, 5, 190, 165)
  ButtonGadget(1, 70, 177, 60, 20, "Play")

  Restore Merlin

  Repeat
    Read item
    If item <> ""
      AddGadgetItem(0, #PB_Any, item)
    EndIf
  Until item = ""

  dhCallMethod(oMerlin, "MoveTo(%d,%d)", WindowX(0) + 40, WindowY(0) - 120)
  dhCallMethod(oMerlin, "Play(%T)", @"GetAttention")
  dhCallMethod(oMerlin, "Play(%T)", @"GetAttentionContinued")
  dhCallMethod(oMerlin, "Speak(%T)", @"Can I do something for you?")
  dhCallMethod(oMerlin, "Play(%T)", @"GetAttentionReturn")
  SetGadgetState(0, 0)

  Repeat

    Select WaitWindowEvent()
      Case #PB_Event_CloseWindow
        CloseWindow(0)
        dhCallMethod(oMerlin, "Stop")
        dhCallMethod(oMerlin, "Play(%T)", @"Hide")
        Delay(3000)
        Break

      Case #PB_Event_MoveWindow
        dhCallMethod(oMerlin, "MoveTo(%d,%d)", WindowX(0) + 40, WindowY(0) - 120)

      Case #PB_Event_Gadget

        Select EventGadget()
          Case 1
            item = GetGadgetText(0)
            If Item = "Speak"
              Item = InputRequester("Merlin", "What should i say?", "")
              If Item <> ""
                dhCallMethod(oMerlin, "Speak(%T)", @Item)
              EndIf
            Else
              dhCallMethod(oMerlin, "Stop")
              dhCallMethod(oMerlin, "Play(%T)", @item)
            EndIf
        EndSelect
    EndSelect

  ForEver
EndIf

dhReleaseObject(oMerlin)
dhReleaseObject(oAgent)

DataSection
  Merlin:
  Data.s "Acknowledge", "Alert", "Announce", "Blink", "Confused", "Congratulate"
  Data.s "Congratulate_2", "Decline", "DoMagic1", "DoMagic2", "DontRecognize"
  Data.s "Explain", "GestureDown", "GestureLeft", "GestureRight", "GestureUp"
  Data.s "GetAttention", "GetAttentionContinued", "GetAttentionReturn", "Greet"
  Data.s "Hearing_1", "Hearing_2", "Hearing_3", "Hearing_4", "Hide", "Idle1_1"
  Data.s "Idle1_2", "Idle1_3", "Idle1_4", "Idle2_1", "Idle2_2", "Idle3_1", "Idle3_2"
  Data.s "LookDown", "LookDownBlink", "LookDownReturn", "LookLeft", "LookLeftBlink"
  Data.s "LookLeftReturn", "LookRight", "LookRightBlink", "LookRightReturn"
  Data.s "LookUp", "LookUpBlink", "LookUpReturn", "MoveDown", "MoveLeft", "MoveRight"
  Data.s "MoveUp", "Pleased", "Process", "Processing", "Read", "ReadContinued"
  Data.s "Reading", "ReadReturn", "Sad", "Search", "Searching", "Show", "Speak"
  Data.s "StartListening", "StopListening", "Suggest", "Surprised", "Think", "Thinking"
  Data.s "Uncertain", "Wave", "Write", "WriteContinued", "WriteReturn", "Writing"
  Data.s ""
EndDataSection

; IDE Options = PureBasic v4.02 (Windows - x86)
; Folding = -
; EnableXP
; HideErrorLog