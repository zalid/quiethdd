; Example by Kiffi

EnableExplicit

Define.l szResponse

Define.l oDom
Define.l oNode
Define.l oNodeList

Define.s XML
Define.s XPath
Define.l CountNodes, NodeCounter

XML = "<adressen>"
XML + " <adresse>"
XML + "  <vorname>Peter</vorname>"
XML + "  <nachname>Parker</nachname>"
XML + " </adresse>"
XML + " <adresse>"
XML + "  <vorname>Bruce</vorname>"
XML + "  <nachname>Wayne</nachname>"
XML + " </adresse>"
XML + " <adresse>"
XML + "  <vorname>Clark</vorname>"
XML + "  <nachname>Kent</nachname>"
XML + " </adresse>"
XML + "</adressen>"

dhToggleExceptions(#True)

oDom = dhCreateObject("MSXML2.DOMDocument.4.0")

If oDom

  dhCallMethod(oDom, "LoadXml(%T)", @XML)

  XPath = "adressen/adresse"

  dhGetValue("%o", @oNodeList, oDom, "SelectNodes(%T)",  @XPath)

  If oNodeList

    dhGetValue("%d", @CountNodes, oNodeList, "length")

    For NodeCounter = 0 To CountNodes - 1

      dhGetValue("%o", @oNode, oNodeList, "item(%d)",  NodeCounter)

      If oNode

        dhGetValue("%T", @szResponse, oNode, "Xml")

        If szResponse
          MessageRequester("XML", PeekS(szResponse))
          dhFreeString(szResponse) : szResponse = 0
        EndIf

      EndIf

    Next

  EndIf

  dhReleaseObject(oNode)      : oNode     = 0
  dhReleaseObject(oNodeList)  : oNodeList = 0
  dhReleaseObject(oDom)       : oDom      = 0

EndIf
