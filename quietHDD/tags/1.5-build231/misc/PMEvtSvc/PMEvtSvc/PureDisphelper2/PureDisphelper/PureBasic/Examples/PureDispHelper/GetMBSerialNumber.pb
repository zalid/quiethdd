; requires puredisphelper
; by ts-soft, based on a vb example

EnableExplicit

Procedure.s MBSerialNumberScript()

  Protected Script$

  Script$ = "Set WMI = GetObject(" + #DQUOTE$ + "WinMgmts:" + #DQUOTE$ + ")" + #CRLF$
  Script$ + "Set objs = WMI.InstancesOf(" + #DQUOTE$ + "Win32_BaseBoard" + #DQUOTE$ + ")" + #CRLF$
  Script$ + "For Each obj in objs" + #CRLF$
  Script$ + "sAns = sAns & obj.SerialNumber" + #CRLF$
  Script$ + "If sAns < objs.Count Then sAns = sAns & "+ #DQUOTE$ + ","+ #DQUOTE$ + #CRLF$
  Script$ + "Next" + #CRLF$
  Script$ + "MBSerialNumber = sAns" + #CRLF$

  ProcedureReturn Script$

EndProcedure

dhToggleExceptions(#True)

Define.l Result, obj = dhCreateObject("MSScriptControl.ScriptControl")
Define.s Script

If obj

  dhPutValue(obj, "Language = %T", @"VBScript")
  Script = MBSerialNumberScript()
  dhCallMethod(obj, "AddCode(%T)", @Script)
  dhGetValue("%T", @Result, obj, "Eval(%T)", @"MBSerialNumber")

  If Result

    MessageRequester("MBSerialNumber:", PeekS(Result), #MB_ICONINFORMATION)
    dhFreeString(Result) : Result = 0

  EndIf

  dhReleaseObject(obj) : obj = 0

EndIf

; IDE Options = PureBasic 4.20 Beta 2 (Windows - x86)
; Folding = -
; EnableXP
; EnableUser
; HideErrorLog