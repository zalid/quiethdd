EnableExplicit

Global oTreeView.l

Enumeration
  #NodeText
EndEnumeration

Enumeration ; LabelEdit
  #tvwAutomatic
  #tvwManual
EndEnumeration

Enumeration ; Relationship
  #tvwFirst
  #tvwLast
  #tvwNext
  #tvwPrevious
  #tvwChild
EndEnumeration

XIncludeFile "..\ComEventSink.pb"

dhToggleExceptions(#True)

Procedure Event_NodeClick()
 
  Protected oTreeNode.l
  Protected szResponse.l

  ; get the selected node
  dhGetValue("%o", @oTreeNode, oTreeView, ".SelectedItem")
 
  If oTreeNode
   
    ; get the text of the selected node
    dhGetValue("%T", @szResponse, oTreeNode, ".Text")
   
    If szResponse
      SetGadgetText(#NodeText, PeekS(szResponse))
      dhFreeString(szResponse) : szResponse = 0
    EndIf
   
    dhReleaseObject(oTreeNode)
   
  EndIf
 
EndProcedure

Procedure EventCallback(Event$, ParameterCount, *Params)
 
  Select Event$
    Case "NodeClick" : Event_NodeClick()
    Case "MouseMove"
    Default          : Debug Event$
  EndSelect
 
EndProcedure

Procedure.l TreeviewNodesAdd(oTreeView.l, Relative.s = "", Relationship.l = -1, key.s = "", Text.s = "")
 
  Protected obj.l
  dhGetValue("%o", @obj, oTreeView, ".Nodes.Add(%T, %d, %T, %T)", @Relative, Relationship, @key, @Text)
  ProcedureReturn obj
 
EndProcedure

Define.l I
Define oTreeNode.l

If OpenWindow(0, #PB_Ignore, #PB_Ignore, 200, 400, "TreeView") And CreateGadgetList(WindowID(0))
 
  StringGadget(#NodeText, 5, 5, 190, 20,"")
 
  oTreeView = OCX_CreateGadget(1, 5, 30, 190, 365, "MSComctlLib.TreeCtrl")
 
  dhPutValue(oTreeView, ".HideSelection=%b", #False)
  dhPutValue(oTreeView, ".LabelEdit=%d", #tvwManual)
  ;dhPutValue(oTreeView, ".Indentation=%d", 5) ; doesn't work. why?

  OCX_ConnectEvents(oTreeView, @EventCallback())
 
  ; ---
 
  ; Add Rootnode
  dhGetValue("%o", @oTreeNode, oTreeView, ".Nodes.Add")
  dhPutValue(oTreeNode, ".Text=%T", @"Root")
  dhPutValue(oTreeNode, ".Key=%T", @"root")
  dhPutValue(oTreeNode, ".Expanded=%b", #True) ; Make sure that every parentnode is expanded (for demo only to see, that there are further subnodes)
 
  ; ---
 
  ; Add First SubNode
  oTreeNode = TreeviewNodesAdd(oTreeView, "root", #tvwChild, "animals", "Animals")
  dhPutValue(oTreeNode, ".Expanded=%b", #True) ; Make sure that every parentnode is expanded (for demo only to see, that there are further subnodes)
 
  ; Add First SubSubNode
  oTreeNode = TreeviewNodesAdd(oTreeView, "animals", #tvwChild, "cats", "Cats")
  ; Add Second SubSubNode
  oTreeNode = TreeviewNodesAdd(oTreeView, "animals", #tvwChild, "dogs", "Dogs")
 
  ; ---
 
  ; Add Second SubNode
  oTreeNode = TreeviewNodesAdd(oTreeView, "root", #tvwChild, "cars", "Cars")
  dhPutValue(oTreeNode, ".Expanded=%b", #True) ; Make sure that every parentnode is expanded (for demo only to see, that there are further subnodes)
 
  ; Add First SubSubNode
  oTreeNode = TreeviewNodesAdd(oTreeView, "cars", #tvwChild, "ferrari", "Ferrari")
  ; Add Second SubSubNode
  oTreeNode = TreeviewNodesAdd(oTreeView, "cars", #tvwChild, "lamborghini", "Lamborghini")
 
  ; ---
 
  ; Now select the first SubSubnode (Cats)
  dhGetValue("%o", @oTreeNode, oTreeView, ".Nodes(%T)", @"cats")
  dhPutValue(oTreeNode, ".Selected=%b", #True)
  Event_NodeClick()
 
  ; ---
 
  Define.l WWE
  Define.s NewNodeText
 
  Repeat
   
    WWE=WaitWindowEvent()
   
    Select WWE
      Case #PB_Event_Gadget
        Select EventGadget()
          Case #NodeText
            If EventType() = #PB_EventType_Change
             
              dhGetValue("%o", @oTreeNode, oTreeView, ".SelectedItem")
             
              If oTreeNode
                NewNodeText=GetGadgetText(#NodeText)
                dhPutValue(oTreeNode, ".Text=%T", @NewNodeText)
                dhReleaseObject(oTreeNode)
              EndIf
             
            EndIf
        EndSelect
    EndSelect
   
  Until WWE = #PB_Event_CloseWindow
 
  CloseWindow(0)
 
  dhReleaseObject(oTreeNode)
  dhReleaseObject(oTreeView)
 
EndIf
; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 20
; Folding = --
; EnableXP
; HideErrorLog