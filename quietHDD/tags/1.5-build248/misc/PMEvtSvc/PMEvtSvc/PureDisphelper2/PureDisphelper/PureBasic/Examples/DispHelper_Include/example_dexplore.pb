; example based on c-source "dexplore.c" in the "samples_c" folder of disphelper_081.zip

; changed to PB by ts-soft

EnableExplicit

XIncludeFile "DispHelper_Include.pb"

; Demonstrates controlling Microsoft's new help system for developers, dexplore.
; This is also known As MS Help 2 And ships With Visual Studio And the platform SDK.


; /* **************************************************************************
;  * ShowHelpTopic:
;  *   Function To show a Platform SDK help topic in dExplore(MS Help 2).
;  *
;  ============================================================================ */

Procedure ShowHelpTopic(szKeyword.s)
;   if you have another version, please change the namespace!
;   some example namespaces:
;   borland.bds3            = Borland Help
;   borland.bds4
;   MS.Dexplore             = Microsoft Dexplore Collection
;   MS.Dexplore.v80.en      = Microsoft Document Explorer Help ENU
;   MS.ENTSERV.v10.en       = Enterprise development And servers except For SQL Server And Exchange
;   MS.KB.v10.en            = Knowledge Base English
;   MS.PSDKSVR2003SP1.1033  = Platform SDK Collection For Windows Server 2003 SP1
;   MS.MSDNQTR.v80.en       = MSDN Library For Visual Studio 2005
;   MS.NETDEV.v10.en        = .NET Development content, ADO.NET, ASP.NET, .NET Fwrk, And MapPoint
;   MS.NETDEVFX.v20.en      = Net Development Framework 2.0
;   MS.SQL.v2005.en         = SDK documentation For SQL Server 2005
;   MS.VisualStudio.v80.en  = Visual Studio 2005
;   MS.WIN32COM.v10.en      = Win32- und COM-Entwicklung

  Protected Namespace.s = "ms-help://MS.PSDKSVR2003SP1.1033"
  
  Protected dExplore.l = dhCreateObject("DExplore.AppObj")
  Protected helpHost.l
  
  If dExplore <> 0
    WITH_(helpHost, dExplore, ".Help")
      dhCallMethod(helpHost, ".SetCollection(%T,%T)", @Namespace, @"Platform SDK")
      dhCallMethod(helpHost, ".SyncIndex(%T,%d)", @szKeyword, 1)
      dhCallMethod(helpHost, ".DisplayTopicFromKeyword(%T)", @szKeyword)
    END_WITH(helpHost)
  EndIf
  ; dexplore 'may' exit when we release the object so we wait
  MessageRequester("dExplore", "Press Ok to close")
  If dExplore <> 0 : dhReleaseObject(dExplore) : EndIf
EndProcedure

dhInitializeImp()
dhToggleExceptions(#True)

ShowHelpTopic("RegCreateKeyEx")

dhUninitialize()
End

; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 35
; FirstLine = 12
; Folding = -
; EnableUnicode
; EnableXP
; HideErrorLog