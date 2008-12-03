; example by Kiffi

EnableExplicit

Define.l oXMLHTTP, oDoc, oNodeList, oNode

Procedure.s GetNodeText(oNode, xPath.s)
  Protected szResponse.l
  Protected ReturnValue.s
  dhGetValue("%T", @szResponse, oNode, "SelectSingleNode(%T).Text", @xPath)
  If szResponse
    ReturnValue = PeekS(szResponse)
    dhFreeString(szResponse) : szResponse = 0
  EndIf
  ProcedureReturn ReturnValue
EndProcedure

oXMLHTTP = dhCreateObject("MSXML2.ServerXMLHTTP")

If oXMLHTTP

  dhCallMethod(oXMLHTTP, ".Open %T, %T, %b", @"GET", @"http://www.codeproject.com/webservices/articlerss.aspx?cat=6", #False)
  dhCallMethod(oXMLHTTP, ".Send")

  dhGetValue("%o", @oDoc, oXMLHTTP, "ResponseXML")

  If oDoc

    dhGetValue("%o", @oNodeList, oDoc, "SelectNodes(%T)",  @"rss/channel/item")

    If oNodeList

      Repeat

        dhGetValue("%o", @oNode, oNodeList, ".NextNode")

        If oNode = 0
          Break
        EndIf

        Debug "subject: "     + GetNodeText(oNode, "subject")
        Debug "title: "       + GetNodeText(oNode, "title")
        Debug "description: " + GetNodeText(oNode, "description")
        Debug "link: "        + GetNodeText(oNode, "link")
        Debug "author: "      + GetNodeText(oNode, "author")
        Debug "category: "    + GetNodeText(oNode, "category")
        Debug "pubDate: "     + GetNodeText(oNode, "pubDate")

        Debug "---------------------------------"

        dhReleaseObject(oNode)     : oNode=0

      ForEver

      dhReleaseObject(oNodeList) : oNodeList=0

    EndIf

    dhReleaseObject(oDoc) : oDoc=0

  EndIf

  dhReleaseObject(oXMLHTTP) : oXMLHTTP = 0

EndIf
