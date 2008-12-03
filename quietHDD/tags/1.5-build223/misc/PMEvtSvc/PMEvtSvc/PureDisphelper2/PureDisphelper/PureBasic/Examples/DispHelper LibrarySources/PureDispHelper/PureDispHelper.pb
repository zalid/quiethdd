EnableExplicit

Import "ole32.lib"
EndImport

Import "uuid.lib"
EndImport

ImportC "disphelper.lib";include
  dhUninitialize(bUninitializeCOM.l = #True) As "_dhUninitialize"
  CompilerIf #PB_Compiler_Unicode
    dhInitializeImp(bInitializeCOM.l = #True, bUnicode.l = #True) As "_dhInitializeImp"
  CompilerElse
    dhInitializeImp(bInitializeCOM.l = #True, bUnicode.l = #False) As "_dhInitializeImp"
  CompilerEndIf
  dhToggleExceptions_(bShow.l) As "_dhToggleExceptions"
  dhGetValue_(szIdentifier.p-unicode, *pResult, *pDisp.IDispatch, szMember.p-bstr, a=0,b=0,c=0,d=0,e=0,f=0,g=0,h=0,i=0,j=0,k=0,l=0) As "_dhGetValue"
  dhPutRef_(*pDisp.IDispatch, szMember.p-bstr, a=0,b=0,c=0,d=0,e=0,f=0,g=0,h=0,i=0,j=0,k=0,l=0) As "_dhPutRef"
  dhPutValue_(*pDisp.IDispatch, szMember.p-bstr, a=0,b=0,c=0,d=0,e=0,f=0,g=0,h=0,i=0,j=0,k=0,l=0) As "_dhPutValue"
  dhCallMethod_(*pDisp.IDispatch, szMember.p-bstr, a=0,b=0,c=0,d=0,e=0,f=0,g=0,h=0,i=0,j=0,k=0,l=0) As "_dhCallMethod"
  dhGetObject_(szFile.p-bstr, szProgId.p-bstr, *ppDisp.IDispatch) As "_dhGetObject"
  dhCreateObject_(szProgId.p-bstr, szMachine.l, *ppDisp.IDispatch) As "_dhCreateObject"
  ;dhEnumBegin_(*ppEnum.IEnumVARIANT, *pDisp.IDispatch, szMember.p-bstr, a=0,b=0,c=0,d=0,e=0,f=0,g=0,h=0,i=0,j=0,k=0,l=0) As "_dhEnumBegin"
  ;dhEnumNextObject_(*pEnum.IEnumVARIANT, *ppDisp.IDispatch) As "_dhEnumNextObject"
EndImport

ProcedureDLL PureDispHelper_Init()
  dhInitializeImp()
EndProcedure

ProcedureDLL PureDispHelper_End()
  dhUninitialize()
EndProcedure

ProcedureDLL dhCreateObject(ProgId.s)
  Protected object.IDispatch, Container.IUnknown
  dhCreateObject_(ProgId, #NUL, @object)
  ProcedureReturn object
EndProcedure

ProcedureDLL dhReleaseObject(obj.l) ; Releases the IDispatch pointer to a COM object
  Protected *object.IUnknown = obj
  If *object
    *object\Release()
  EndIf
EndProcedure

ProcedureDLL dhFreeString(lpString.l); Frees the memory that contains a string returned from dhGetValue
  SysFreeString_(lpString)
EndProcedure

ProcedureDLL dhToggleExceptions(bTrueFalse.l)
  dhToggleExceptions_(bTrueFalse)
EndProcedure

ProcedureDLL dhCallMethod13(obj.l, Methode.s, a, b, c, d, e, f, g, h, i, j, k, l)
  ProcedureReturn dhCallMethod_(obj.l, Methode.s, a, b, c, d, e, f, g, h, i, j, k, l)
EndProcedure
ProcedureDLL dhCallMethod12(obj.l, Methode.s, a, b, c, d, e, f, g, h, i, j, k)
  ProcedureReturn dhCallMethod_(obj.l, Methode.s, a, b, c, d, e, f, g, h, i, j, k)
EndProcedure
ProcedureDLL dhCallMethod11(obj.l, Methode.s, a, b, c, d, e, f, g, h, i, j)
  ProcedureReturn dhCallMethod_(obj.l, Methode.s, a, b, c, d, e, f, g, h, i, j)
EndProcedure
ProcedureDLL dhCallMethod10(obj.l, Methode.s, a, b, c, d, e, f, g, h, i)
  ProcedureReturn dhCallMethod_(obj.l, Methode.s, a, b, c, d, e, f, g, h, i)
EndProcedure
ProcedureDLL dhCallMethod9(obj.l, Methode.s, a, b, c, d, e, f, g, h)
  ProcedureReturn dhCallMethod_(obj.l, Methode.s, a, b, c, d, e, f, g, h)
EndProcedure
ProcedureDLL dhCallMethod8(obj.l, Methode.s, a, b, c, d, e, f, g)
  ProcedureReturn dhCallMethod_(obj.l, Methode.s, a, b, c, d, e, f, g)
EndProcedure
ProcedureDLL dhCallMethod7(obj.l, Methode.s, a, b, c, d, e, f)
  ProcedureReturn dhCallMethod_(obj.l, Methode.s, a, b, c, d, e, f)
EndProcedure
ProcedureDLL dhCallMethod6(obj.l, Methode.s, a, b, c, d, e)
  ProcedureReturn dhCallMethod_(obj.l, Methode.s, a, b, c, d, e)
EndProcedure
ProcedureDLL dhCallMethod5(obj.l, Methode.s, a, b, c, d)
  ProcedureReturn dhCallMethod_(obj.l, Methode.s, a, b, c, d)
EndProcedure
ProcedureDLL dhCallMethod4(obj.l, Methode.s, a, b, c)
  ProcedureReturn dhCallMethod_(obj.l, Methode.s, a, b, c)
EndProcedure
ProcedureDLL dhCallMethod3(obj.l, Methode.s, a, b)
  ProcedureReturn dhCallMethod_(obj.l, Methode.s, a, b)
EndProcedure
ProcedureDLL dhCallMethod2(obj.l, Methode.s, a)
  ProcedureReturn dhCallMethod_(obj.l, Methode.s, a)
EndProcedure
ProcedureDLL dhCallMethod(obj.l, Methode.s) ; Calls a method in a COM object
  ProcedureReturn dhCallMethod_(obj.l, Methode.s)
EndProcedure

ProcedureDLL dhGetObject(File.s, ProgID.s, *IID)
  ProcedureReturn dhGetObject_(File, ProgID, *IID)
EndProcedure

ProcedureDLL dhGetValue13(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c,d,e,f,g,h,i,j,k,l)
  ProcedureReturn dhGetValue_(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c,d,e,f,g,h,i,j,k,l)
EndProcedure
ProcedureDLL dhGetValue12(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c,d,e,f,g,h,i,j,k)
  ProcedureReturn dhGetValue_(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c,d,e,f,g,h,i,j,k)
EndProcedure
ProcedureDLL dhGetValue11(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c,d,e,f,g,h,i,j)
  ProcedureReturn dhGetValue_(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c,d,e,f,g,h,i,j)
EndProcedure
ProcedureDLL dhGetValue10(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c,d,e,f,g,h,i)
  ProcedureReturn dhGetValue_(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c,d,e,f,g,h,i)
EndProcedure
ProcedureDLL dhGetValue9(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c,d,e,f,g,h)
  ProcedureReturn dhGetValue_(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c,d,e,f,g,h)
EndProcedure
ProcedureDLL dhGetValue8(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c,d,e,f,g)
  ProcedureReturn dhGetValue_(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c,d,e,f,g)
EndProcedure
ProcedureDLL dhGetValue7(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c,d,e,f)
  ProcedureReturn dhGetValue_(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c,d,e,f)
EndProcedure
ProcedureDLL dhGetValue6(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c,d,e)
  ProcedureReturn dhGetValue_(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c,d,e)
EndProcedure
ProcedureDLL dhGetValue5(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c,d)
  ProcedureReturn dhGetValue_(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c,d)
EndProcedure
ProcedureDLL dhGetValue4(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c)
  ProcedureReturn dhGetValue_(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b,c)
EndProcedure
ProcedureDLL dhGetValue3(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b)
  ProcedureReturn dhGetValue_(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a,b)
EndProcedure
ProcedureDLL dhGetValue2(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a)
  ProcedureReturn dhGetValue_(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s, a)
EndProcedure
ProcedureDLL dhGetValue(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s) ; Returns a value from a function or property of a COM object
  ProcedureReturn dhGetValue_(formatdentifier.s, *RetVal, obj.l, FunctionOrProperty.s)
EndProcedure

ProcedureDLL dhPutRef13(obj.l, Property.s, a, b, c, d, e, f, g, h, i, j, k, l)
  ProcedureReturn dhPutRef_(obj, Property, a, b, c, d, e, f, g, h, i, j, k, l)
EndProcedure
ProcedureDLL dhPutRef12(obj.l, Property.s, a, b, c, d, e, f, g, h, i, j, k)
  ProcedureReturn dhPutRef_(obj, Property, a, b, c, d, e, f, g, h, i, j, k)
EndProcedure
ProcedureDLL dhPutRef11(obj.l, Property.s, a, b, c, d, e, f, g, h, i, j)
  ProcedureReturn dhPutRef_(obj, Property, a, b, c, d, e, f, g, h, i, j)
EndProcedure
ProcedureDLL dhPutRef10(obj.l, Property.s, a, b, c, d, e, f, g, h, i)
  ProcedureReturn dhPutRef_(obj, Property, a, b, c, d, e, f, g, h, i)
EndProcedure
ProcedureDLL dhPutRef9(obj.l, Property.s, a, b, c, d, e, f, g, h)
  ProcedureReturn dhPutRef_(obj, Property, a, b, c, d, e, f, g, h)
EndProcedure
ProcedureDLL dhPutRef8(obj.l, Property.s, a, b, c, d, e, f, g)
  ProcedureReturn dhPutRef_(obj, Property, a, b, c, d, e, f, g)
EndProcedure
ProcedureDLL dhPutRef7(obj.l, Property.s, a, b, c, d, e, f)
  ProcedureReturn dhPutRef_(obj, Property, a, b, c, d, e, f)
EndProcedure
ProcedureDLL dhPutRef6(obj.l, Property.s, a, b, c, d, e)
  ProcedureReturn dhPutRef_(obj, Property, a, b, c, d, e)
EndProcedure
ProcedureDLL dhPutRef5(obj.l, Property.s, a, b, c, d)
  ProcedureReturn dhPutRef_(obj, Property, a, b, c, d)
EndProcedure
ProcedureDLL dhPutRef4(obj.l, Property.s, a, b, c)
  ProcedureReturn dhPutRef_(obj, Property, a, b, c)
EndProcedure
ProcedureDLL dhPutRef3(obj.l, Property.s, a, b)
  ProcedureReturn dhPutRef_(obj, Property, a, b)
EndProcedure
ProcedureDLL dhPutRef2(obj.l, Property.s, a)
  ProcedureReturn dhPutRef_(obj, Property, a)
EndProcedure
ProcedureDLL dhPutRef(obj.l, Property.s) ; Sets a property of a COM object to an object reference
  ProcedureReturn dhPutRef_(obj, Property)
EndProcedure

ProcedureDLL dhPutValue12(obj.l, Property.s, a, b, c, d, e, f, g, h, i, j, k, l)
  ProcedureReturn dhPutValue_(obj, Property, a, b, c, d, e, f, g, h, i, j, k, l)
EndProcedure
ProcedureDLL dhPutValue11(obj.l, Property.s, a, b, c, d, e, f, g, h, i, j, k)
  ProcedureReturn dhPutValue_(obj, Property, a, b, c, d, e, f, g, h, i, j, k)
EndProcedure
ProcedureDLL dhPutValue10(obj.l, Property.s, a, b, c, d, e, f, g, h, i, j)
  ProcedureReturn dhPutValue_(obj, Property, a, b, c, d, e, f, g, h, i, j)
EndProcedure
ProcedureDLL dhPutValue9(obj.l, Property.s, a, b, c, d, e, f, g, h, i)
  ProcedureReturn dhPutValue_(obj, Property, a, b, c, d, e, f, g, h, i)
EndProcedure
ProcedureDLL dhPutValue8(obj.l, Property.s, a, b, c, d, e, f, g, h)
  ProcedureReturn dhPutValue_(obj, Property, a, b, c, d, e, f, g, h)
EndProcedure
ProcedureDLL dhPutValue7(obj.l, Property.s, a, b, c, d, e, f, g)
  ProcedureReturn dhPutValue_(obj, Property, a, b, c, d, e, f, g)
EndProcedure
ProcedureDLL dhPutValue6(obj.l, Property.s, a, b, c, d, e, f)
  ProcedureReturn dhPutValue_(obj, Property, a, b, c, d, e, f)
EndProcedure
ProcedureDLL dhPutValue5(obj.l, Property.s, a, b, c, d, e)
  ProcedureReturn dhPutValue_(obj, Property, a, b, c, d, e)
EndProcedure
ProcedureDLL dhPutValue4(obj.l, Property.s, a, b, c, d)
  ProcedureReturn dhPutValue_(obj, Property, a, b, c, d)
EndProcedure
ProcedureDLL dhPutValue3(obj.l, Property.s, a, b, c)
  ProcedureReturn dhPutValue_(obj, Property, a, b, c)
EndProcedure
ProcedureDLL dhPutValue2(obj.l, Property.s, a, b)
  ProcedureReturn dhPutValue_(obj, Property, a, b)
EndProcedure
ProcedureDLL dhPutValue(obj.l, Property.s, a) ; Sets a property value of a COM object
  ProcedureReturn dhPutValue_(obj, Property, a)
EndProcedure

; ProcedureDLL dhEnumBegin13(*ppEnum, *pDisp, szMember.s, a, b, c, d, e, f, g, h, i, j, k, l)
;   ProcedureReturn dhEnumBegin_(*ppEnum, *pDisp, szMember.s, a, b, c, d, e, f, g, h, i, j, k, l)
; EndProcedure
; ProcedureDLL dhEnumBegin12(*ppEnum, *pDisp, szMember.s, a, b, c, d, e, f, g, h, i, j, k)
;   ProcedureReturn dhEnumBegin_(*ppEnum, *pDisp, szMember.s, a, b, c, d, e, f, g, h, i, j, k)
; EndProcedure
; ProcedureDLL dhEnumBegin11(*ppEnum, *pDisp, szMember.s, a, b, c, d, e, f, g, h, i, j)
;   ProcedureReturn dhEnumBegin_(*ppEnum, *pDisp, szMember.s, a, b, c, d, e, f, g, h, i, j)
; EndProcedure
; ProcedureDLL dhEnumBegin10(*ppEnum, *pDisp, szMember.s, a, b, c, d, e, f, g, h, i)
;   ProcedureReturn dhEnumBegin_(*ppEnum, *pDisp, szMember.s, a, b, c, d, e, f, g, h, i)
; EndProcedure
; ProcedureDLL dhEnumBegin9(*ppEnum, *pDisp, szMember.s, a, b, c, d, e, f, g, h)
;   ProcedureReturn dhEnumBegin_(*ppEnum, *pDisp, szMember.s, a, b, c, d, e, f, g, h)
; EndProcedure
; ProcedureDLL dhEnumBegin8(*ppEnum, *pDisp, szMember.s, a, b, c, d, e, f, g)
;   ProcedureReturn dhEnumBegin_(*ppEnum, *pDisp, szMember.s, a, b, c, d, e, f, g)
; EndProcedure
; ProcedureDLL dhEnumBegin7(*ppEnum, *pDisp, szMember.s, a, b, c, d, e, f)
;   ProcedureReturn dhEnumBegin_(*ppEnum, *pDisp, szMember.s, a, b, c, d, e, f)
; EndProcedure
; ProcedureDLL dhEnumBegin6(*ppEnum, *pDisp, szMember.s, a, b, c, d, e)
;   ProcedureReturn dhEnumBegin_(*ppEnum, *pDisp, szMember.s, a, b, c, d, e)
; EndProcedure
; ProcedureDLL dhEnumBegin5(*ppEnum, *pDisp, szMember.s, a, b, c, d)
;   ProcedureReturn dhEnumBegin_(*ppEnum, *pDisp, szMember.s, a, b, c, d)
; EndProcedure
; ProcedureDLL dhEnumBegin4(*ppEnum, *pDisp, szMember.s, a, b, c)
;   ProcedureReturn dhEnumBegin_(*ppEnum, *pDisp, szMember.s, a, b, c)
; EndProcedure
; ProcedureDLL dhEnumBegin3(*ppEnum, *pDisp, szMember.s, a, b)
;   ProcedureReturn dhEnumBegin_(*ppEnum, *pDisp, szMember.s, a, b)
; EndProcedure
; ProcedureDLL dhEnumBegin2(*ppEnum, *pDisp, szMember.s, a)
;   ProcedureReturn dhEnumBegin_(*ppEnum, *pDisp, szMember.s, a)
; EndProcedure
; ProcedureDLL dhEnumBegin(*ppEnum, *pDisp, szMember.s)
;   ProcedureReturn dhEnumBegin_(*ppEnum, *pDisp, szMember.s)
; EndProcedure
; 
; ProcedureDLL dhEnumNextObject(*pEnum, *ppDisp)
;   ProcedureReturn dhEnumNextObject_(*pEnum, *ppDisp)
; EndProcedure


; IDE Options = PureBasic 4.20 Beta 2 (Windows - x86)
; CursorPosition = 5
; Folding = -AAAAAAAAA9
; HideErrorLog