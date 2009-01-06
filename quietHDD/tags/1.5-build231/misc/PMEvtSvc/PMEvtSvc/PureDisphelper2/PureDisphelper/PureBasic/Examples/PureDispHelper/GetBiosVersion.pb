EnableExplicit

Procedure.s GetBiosVersion()

  Protected Script$

  Script$ = "Set objWMIService = GetObject(" + #DQUOTE$ + "winmgmts:\\.\root\CIMV2" + #DQUOTE$ + ")" + #CRLF$
  Script$ + "Set colItems = objWMIService.ExecQuery(" + #DQUOTE$ + "SELECT * FROM Win32_BIOS" + #DQUOTE$ + ",,48)" + #CRLF$
  Script$ + "For Each objItem in colItems" + #CRLF$
  Script$ + "If isNull(objItem.BIOSVersion) Then" + #CRLF$
  Script$ + "myBios = I'm sure, I don't know" + #CRLF$
  Script$ + "Else" + #CRLF$
  Script$ + "myBios = Join(objItem.BIOSVersion, " + #DQUOTE$ + "," + #DQUOTE$ + ")" + #CRLF$
  Script$ + "End If" + #CRLF$
  Script$ + "Next"

  ProcedureReturn Script$

EndProcedure

dhToggleExceptions(#True)

Define.l Result, obj = dhCreateObject("MSScriptControl.ScriptControl")
Define.s Script

If obj

  dhPutValue(obj, "Language = %T", @"VBScript")
  Script = GetBiosVersion()
  dhCallMethod(obj, "AddCode(%T)", @Script)
  dhGetValue("%T", @Result, obj, "Eval(%T)", @"myBios")

  If Result

    MessageRequester("BIOSVersion:", PeekS(Result), #MB_ICONINFORMATION)
    dhFreeString(Result) : Result = 0

  EndIf

  dhReleaseObject(obj) : obj = 0

EndIf

; IDE Options = PureBasic 4.20 Beta 2 (Windows - x86)
; Folding = -
; EnableXP
; EnableUser
; HideErrorLog