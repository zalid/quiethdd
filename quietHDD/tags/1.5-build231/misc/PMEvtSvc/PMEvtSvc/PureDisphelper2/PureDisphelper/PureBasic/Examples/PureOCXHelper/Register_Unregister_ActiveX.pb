Procedure RegisterActiveX(DLLName.s)
  Protected DLL.l, Result.l
  CoInitialize_(0)
  DLL = OpenLibrary(#PB_Any, DLLName)
  If DLL
    If CallFunction(DLL, "DllRegisterServer") = #S_OK
      Result = #True
    EndIf
    CloseLibrary(DLL)
  EndIf
  ProcedureReturn Result
EndProcedure

Procedure UnRegisterActiveX(DLLName.s)
  Protected DLL.l, Result.l
  CoInitialize_(0)
  DLL = OpenLibrary(#PB_Any, DLLName)
  If DLL
    If CallFunction(DLL, "DllUnregisterServer") = #S_OK
      Result = #True
    EndIf
    CloseLibrary(DLL)
  EndIf
  ProcedureReturn Result
EndProcedure
; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 24
; Folding = -
; EnableXP
; HideErrorLog