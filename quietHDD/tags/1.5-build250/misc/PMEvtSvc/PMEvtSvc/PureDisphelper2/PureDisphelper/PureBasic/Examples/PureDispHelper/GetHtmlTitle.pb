; example by ts-soft

EnableExplicit

Define.l oIE, oDoc, Result

dhToggleExceptions(#True)
oIE = dhCreateObject("InternetExplorer.Application")

If oIE

  dhPutValue(oIE, "Visible = %b", #False)
  dhCallMethod(oIE, "Navigate (%T)", @"www.ts-soft.eu")

  Repeat

    dhGetValue("%d", @Result, oIE, "ReadyState")

  Until Result = 4

  dhGetValue("%o", @oDoc, oIE, "Document")

  If oDoc

    dhGetValue("%T", @Result, oDoc, "title")

    If Result

      MessageRequester("Title from www.ts-soft.eu", PeekS(Result), #MB_ICONINFORMATION)
      dhFreeString(Result) : Result = 0

    EndIf

    dhReleaseObject(oDoc)

  EndIf

  dhReleaseObject(oIE)

EndIf
