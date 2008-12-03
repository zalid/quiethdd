; by Kiffi and ts-soft

EnableExplicit

Procedure.s GetHtmlTitleVbs(URL$)

  Protected Script$

  Script$ = "Set myIE = CreateObject(" + #DQUOTE$ + "InternetExplorer.Application" + #DQUOTE$ + ")" + #CRLF$
  Script$ + "Do While myIE.Busy" + #CRLF$
  Script$ + "Loop" + #CRLF$
  Script$ + "myIE.Visible = False" + #CRLF$
  Script$ + "myIE.Navigate " + #DQUOTE$ + URL$ + #DQUOTE$ + #CRLF$
  Script$ + "Do While myIE.ReadyState <> 4" + #CRLF$
  Script$ + "Loop" + #CRLF$
  Script$ + "Set oDoc = myIE.Document" + #CRLF$
  Script$ + "myTitle = oDoc.title" + #CRLF$
  Script$ + "Set oDoc = Nothing" + #CRLF$
  Script$ + "Set myIE = Nothing" + #CRLF$

  ProcedureReturn Script$

EndProcedure

dhToggleExceptions(#True)

Define.l Result, obj = dhCreateObject("MSScriptControl.ScriptControl")
Define.s Script

If obj

  dhPutValue(obj, "Language = %T", @"VBScript")
  dhGetValue("%T", @Result, obj, "Language")

  If Result

    Debug "Language: " + PeekS(Result)
    dhFreeString(Result) : Result = 0

  EndIf

  dhPutValue(obj, "TimeOut = %d", 20000)
  dhGetValue("%d", @Result, obj, "TimeOut")

  Debug "TimeOut: " + Str(Result) + " ms"
  Script = GetHtmlTitleVbs("www.purebasic.com")
  dhCallMethod(obj, "AddCode(%T)", @Script)
  dhGetValue("%T", @Result, obj, "Eval(%T)", @"myTitle")

  If Result

    MessageRequester("PureBasic.com Titel:", PeekS(Result), #MB_ICONINFORMATION)
    dhFreeString(Result) : Result = 0

  EndIf

  dhReleaseObject(obj) : obj = 0

EndIf
