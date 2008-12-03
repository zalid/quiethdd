; based on FreeBASIC example
; changed and extended to PB by ts-soft

EnableExplicit

dhToggleExceptions(#True)

Define.l WinFlag, szResponse, obj = dhCreateObject("MSXML2.XMLHTTP.4.0")

WinFlag = #PB_Window_SystemMenu | #PB_Window_SizeGadget | #PB_Window_MaximizeGadget | #PB_Window_MinimizeGadget

If ExamineDesktops() And DesktopWidth(0) <= 1024

    WinFlag | #PB_Window_Maximize

EndIf

If OpenWindow(0, #PB_Ignore, #PB_Ignore, 1024, 768, "Please wait...", WinFlag)

  CreateGadgetList(WindowID(0))
  EditorGadget(0, 0, 0, 0, 0, #PB_Editor_ReadOnly)

  If obj

    dhCallMethod(obj, "Open(%T, %T, %b)", @"GET", @"http://sourceforge.net", #False)
    dhCallMethod(obj, "Send")
    dhGetValue("%T", @szResponse, obj, "ResponseText")

    If szResponse

      SetGadgetText(0, PeekS(szResponse))
      SetWindowTitle(0, "http://sourceforge.net")
      dhFreeString(szResponse) : szResponse = 0

    EndIf

    dhReleaseObject(obj) : obj = 0

  EndIf

  Repeat

    Select WaitWindowEvent()

      Case #PB_Event_CloseWindow
        Break

      Case #PB_Event_SizeWindow
        ResizeGadget(0, #PB_Ignore, #PB_Ignore, WindowWidth(0), WindowHeight(0))

    EndSelect

  ForEver

EndIf
