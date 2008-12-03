; Example by Kiffi

EnableExplicit

XIncludeFile "DispHelper_Include.pb"

Define.s XML

XML = "<addresses>"
XML + " <address>"
XML + "  <forename>Peter</forename>"
XML + "  <surname>Parker</surname>"
XML + " </address>"
XML + " <address>"
XML + "  <forename>Bruce</forename>"
XML + "  <surname>Wayne</surname>"
XML + " </address>"
XML + " <address>"
XML + "  <forename>Clark</forename>"
XML + "  <surname>Kent</surname>"
XML + " </address>"
XML + "</addresses>"

dhInitializeImp()

dhToggleExceptions(#True)

Define.l oDom
Define.l oNode

Define.l szResponse

oDom = dhCreateObject("MSXML.DOMDocument")

If oDom

 dhCallMethod(oDom, "LoadXml(%T)", @XML)

 FOR_EACH(oNode, oDom, "SelectNodes(%T)",  @"addresses/address")

   dhGetValue("%T", @szResponse, oNode, "Xml")

   If szResponse
     MessageRequester("XML", PeekS(szResponse))
     dhFreeString(szResponse)
   EndIf

 NEXT_(oNode)
 dhReleaseObject(oDom)
EndIf

dhUninitialize()

; IDE Options = PureBasic 4.20 Beta 2 (Windows - x86)
; CursorPosition = 49
; Folding = -
; EnableXP
; HideErrorLog