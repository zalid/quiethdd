; example by KIKI

EnableExplicit

 ;'Base object for working with Ooo
Declare MakePropertyValue(cName.s, uValue.b)
Declare MakePropertyValue1(cName.s, uValue.s)
Declare.s ConvertToUrl(strFile.s)

Define DISPHELPER_NO_OCX_CreateGadget
Define DISPHELPER_NO_FOR_EACH
Define ocurseur,mdoc,chaine.s
Define.l oSM, oDesk, oDoc

XIncludeFile  "DispHelper_Include.pb"
XIncludeFile "VariantHelper_Include.pb"
Define.safearray *openpar
Define.VARIANT openarray

*openpar = saCreateSafeArray(#VT_DISPATCH, 0, 3)
SA_DISPATCH(*openpar, 0) = MakePropertyValue("ReadOnly", #True)
SA_DISPATCH(*openpar, 1) = MakePropertyValue1("Password", "puredisplay")
SA_DISPATCH(*openpar, 2) = MakePropertyValue("Hidden", #False)
V_ARRAY_DISP(openarray) = *openpar
; Initialization
dhInitializeImp()
dhToggleExceptions(#True)



oSM = dhCreateObject("com.sun.star.ServiceManager")
If oSM
  ;Creating instance of Desktop
  dhGetValue("%o",@oDesk,oSM, ".CreateInstance(%T)", @"com.sun.star.frame.Desktop")
  ;Opening a new writer Document
  dhGetValue("%o",@oDoc, oDesk,".LoadComponentFromURL(%T,%T,%d,%v)",@"private:factory/swriter", @"_blank", 0, openarray)
  ;Getting Text of the document
  dhGetValue("%o",@mdoc, oDoc,".GetText(")
  ;Creating a cursor
  dhGetValue("%o",@ocurseur, mdoc,"createTextCursor(")
  ; Writing text with help of cursor
  chaine.s = "That s a doc created with Help of Pure Display Helper"+ #LF$
  dhCallMethod(mdoc,".insertString(%o, %T, %b",ocurseur, @chaine, #False)
  ;Going to th end of document
  dhCallMethod(ocurseur,".gotoEnd(%b", #False)
  ;And then writing others thiongs
  dhCallMethod(mdoc,".insertString(%o, %T, %b",ocurseur,@"From TS-SOFT To see what we can do with these lib" ,#False)
  MessageRequester("PureDispHelper-OpenWriterDemo", "Click OK to close Open Office Writer") ;
  ; Then close document and releasing object
  dhCallMethod(oDoc,".close(%b)",#True)
  dhReleaseObject(mdoc)
  dhReleaseObject(oDoc)
  dhReleaseObject(oSM)
EndIf
 
 
dhUninitialize()

End

Procedure MakePropertyValue(cName.s, uValue.b)
  Define  oStruct.l
  Define oServiceManager.l = dhCreateObject("com.sun.star.ServiceManager")
  dhGetValue("%o", @oStruct, oServiceManager, ".Bridge_GetStruct(%T)", @"com.sun.star.beans.PropertyValue")
  dhPutValue(oStruct, ".Name=%T", @cName)
  dhPutValue(oStruct, ".Value=%b", @uValue)
  ProcedureReturn oStruct
EndProcedure


Procedure MakePropertyValue1(cName.s, uValue.s)
  Define oServiceManager.l
  Define oStruct
  oServiceManager = dhCreateObject("com.sun.star.ServiceManager")
  dhGetValue("%o", @oStruct, oServiceManager, ".Bridge_GetStruct(%T)", @"com.sun.star.beans.PropertyValue")
  dhPutValue(oStruct, ".Name=%T", @cName)
  dhPutValue(oStruct, ".Value=%T", @uValue)
  ProcedureReturn oStruct
EndProcedure

Procedure.s ConvertToUrl(strFile.s)
  strFile = ReplaceString(strFile, "\", "/")
  strFile = ReplaceString(strFile, ":", "|")
  strFile= ReplaceString(strFile, " ", "%20")
  strFile = "file:///" + strFile
  ProcedureReturn strFile
EndProcedure 
; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 46
; FirstLine = 18
; Folding = -
; EnableXP
; HideErrorLog