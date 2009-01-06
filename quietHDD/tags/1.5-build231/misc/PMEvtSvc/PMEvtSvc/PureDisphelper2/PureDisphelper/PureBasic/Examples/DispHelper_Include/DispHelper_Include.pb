
; Import and Include for Disphelper for PureBasic
; Autor: Thomas (ts-soft) Schulz

; you can comment the following line out or define this variable in your source
; before you include the DispHelper_Include File, if you don't use it

;Define DISPHELPER_NO_FOR_EACH
;Define DISPHELPER_NO_OCX_CreateGadget
;Define DISPHELPER_NO_WITH


Import "atl.lib"
  AtlAxCreateControl(lpszName.p-bstr,hWnd.l,*pStream.IStream,*ppUnkContainer.IUnknown) As "_AtlAxCreateControl"
  AtlAxGetControl(hWnd.l,*pp.IUnknown) As "_AtlAxGetControl"
  AtlAxWinInit() As "_AtlAxWinInit"
EndImport

Import "ole32.lib" : EndImport
Import "oleaut32.lib" : EndImport
Import "uuid.lib" : EndImport

Structure DH_EXCEPTION
  szInitialFunction.s
  szErrorFunction.s
  hr.l
  szMember.c[64]
  szCompleteMember.c[255]
  swCode.l
  szDescription.s
  szSource.s
  szHelpFile.s
  dwHelpContext.l
  iArgError.l
  bDispatchError.l
EndStructure

PrototypeC DH_EXCEPTION_CALLBACK(*PDH_EXCEPTION.DH_EXCEPTION)

Structure DH_EXCEPTION_OPTIONS
  hWnd.l
  szAppName.s
  bShowExceptions.l
  bDisableRecordExceptions.l
  pfnExeptionCallback.DH_EXCEPTION_CALLBACK
EndStructure

ImportC "disphelper.lib"
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
  dhGetObject_(szFile, szProgId, *ppDisp.IDispatch) As "_dhGetObject"
  dhCreateObject_(szProgId.p-bstr, szMachine.l, *ppDisp.IDispatch) As "_dhCreateObject"
  dhEnumBegin(*ppEnum.IEnumVARIANT, *pDisp.IDispatch, szMember.p-bstr, a=0,b=0,c=0,d=0,e=0,f=0,g=0,h=0,i=0,j=0,k=0,l=0) As "_dhEnumBegin"
  dhEnumNextObject(*pEnum.IEnumVARIANT, *ppDisp.IDispatch) As "_dhEnumNextObject"
  dhSetExceptionOptions(*pExceptionOptions.DH_EXCEPTION_OPTIONS) As "_dhSetExceptionOptions"
  dhGetExceptionOptions(*pExceptionOptions.DH_EXCEPTION_OPTIONS) As "_dhGetExceptionOptions"
  dhShowException(*pException.DH_EXCEPTION) As "_dhShowException"
  dhGetLastException(*pException.DH_EXCEPTION) As "_dhGetLastException"
  CompilerIf #PB_Compiler_Unicode
    dhFormatException(*pException.DH_EXCEPTION, szBuffer.s, cchBufferSize.l, bFixedFont.l) As "_dhFormatExceptionW"
  CompilerElse
    dhFormatException(*pException.DH_EXCEPTION, szBuffer.s, cchBufferSize.l, bFixedFont.l) As "_dhFormatExceptionA"
  CompilerEndIf
EndImport

;/////////////////////////////////////////////////////////////////////////////////
;Added by srod to fix a problem with dhgetobject().
  Procedure dhGetObject_Internal(a, b, *c,iDispatch)
    Protected p1, p2
    If PeekS(a, -1, #PB_Unicode)
      p1 = a
    EndIf
    If PeekS(b, -1, #PB_Unicode)
      p2 = b
    EndIf
    dhGetObject_(p1, p2, *c)
  EndProcedure
 
  Prototype.l dh_dhgetobjectProto(a.p-bstr, b.p-bstr, *c.IDispatch)
    Global dh_dhgetobjectProto.dh_dhgetobjectProto = @dhGetObject_Internal()
;/////////////////////////////////////////////////////////////////////////////////

Global AtlAXIsInit.l

CompilerIf Defined(TYPEATTR, #PB_Structure) = #False
  Structure TYPEATTR
    guid.GUID
    lcid.l
    dwReserved.l
    memidConstructor.l
    memidDestructor.l
    lpstrSchema.l
    cbSizeInstance.l
    typekind.l
    cFuncs.w
    cVars.w
    cImplTypes.w
    cbSizeVft.w
    cbAlignment.w
    wTypeFlags.w
    wMajorVerNum.w
    wMinorVerNum.w
    tdescAlias.l
    idldescType.l
  EndStructure
CompilerEndIf

Enumeration
  #TKIND_ENUM
  #TKIND_RECORD
  #TKIND_MODULE
  #TKIND_INTERFACE
  #TKIND_DISPATCH
  #TKIND_COCLASS
  #TKIND_ALIAS
  #TKIND_UNION
  #TKIND_MAX
EndEnumeration

CompilerIf Defined(StringToBStr, #PB_Procedure) = #False
  Procedure StringToBStr(string.s)
    Protected *Buffer = AllocateMemory(StringByteLength(string, #PB_Unicode) + 2)
    Protected bstr_string.l
    If *Buffer
      PokeS(*Buffer, string, -1, #PB_Unicode)
      bstr_string = SysAllocString_(*Buffer)
      FreeMemory(*Buffer)
    EndIf
    ProcedureReturn bstr_string
  EndProcedure
CompilerEndIf
CompilerIf Defined(DISPHELPER_NO_OCX_CreateGadget, #PB_Variable) = #False
  Procedure OCX_CreateGadget_(Gadget.l, x.l, y.l, width.l, height.l, ProgID.s, CLSID.s = "")
    Protected CLSIDGUID.GUID, Mem.l, ProgID_u.l, Container.IUnknown, Object.IDispatch
    Protected pLibPath.l, hKey.l, lpcbData = 255, lpData.s{255}, lType.l = #REG_SZ
    Protected oTypeLib.ITypeLib, Err.l, I.l, J.l, typekind.l, oITypeInfo.ITypeInfo, aTypeAttributes.l
    Protected *oTypeAttributes.TYPEATTR, ptrIID.l, lpData_u.l, CurrentIID.s, hRefType.l, oISubTypeInfo.ITypeInfo
    Protected aSubTypeAttributes.l, *oSubTypeAttributes.TYPEATTR, ptrSubIID.l, MainIID.s, ptrMainIID.l
    Protected MainIID_u.l, ppTInfo.ITypeInfo, pTAttr.l, *oTypeAttr.TYPEATTR, IIDwStr.l
 
    ; we use CLSID
    If CLSID = ""
      ProgID_u = StringToBStr(ProgID)
      CLSIDFromProgID_(ProgID_u, @CLSIDGUID)
      SysFreeString_(ProgID_u)
      StringFromCLSID_(@CLSIDGUID, @Mem)
      ProgID = PeekS(Mem, #PB_Any, #PB_Unicode)
      SysFreeString_(Mem)
    Else
      ProgID = CLSID
    EndIf
 
    ; create a container
    If Gadget = #PB_Any
      Gadget = ContainerGadget(#PB_Any, x, y, width,    height)
    Else
      ContainerGadget(Gadget, x, y, width,    height)
    EndIf
    CloseGadgetList()
 
    If AtlAXIsInit = #False
      AtlAxWinInit()
      AtlAXIsInit = #True
    EndIf
 
    AtlAxCreateControl(ProgID, GadgetID(Gadget), 0, @Container)
    If Container
      AtlAxGetControl(GadgetID(Gadget), @Object)
      ; search for right Dispatch or Interface pointer and some testing stuff
      If Object
        Err = RegOpenKeyEx_(#HKEY_CLASSES_ROOT, "CLSID\" + ProgID + "\InprocServer32", 0, #KEY_ALL_ACCESS, @hKey)
        If Err = #ERROR_SUCCESS
          Err = RegQueryValueEx_(hKey, "", 0, @lType, @lpData, @lpcbData)
          If Err = #ERROR_SUCCESS
            ; we have the path to the ocx and can load it
            lpData_u = StringToBStr(lpData)
            Err = LoadTypeLibEx_(lpData_u, 2, @oTypeLib)
            SysFreeString_(lpData_u)
            If Err = #S_OK
              For I = 0 To oTypeLib\GetTypeInfoCount() - 1
                oTypeLib\GetTypeInfoType(I, @typekind)
                oTypeLib\GetTypeInfo(I, @oITypeInfo)
                oITypeInfo\GetTypeAttr(@aTypeAttributes)
                *oTypeAttributes = aTypeAttributes
                StringFromCLSID_(*oTypeAttributes\guid, @ptrIID)
                CurrentIID = PeekS(ptrIID, #PB_Any, #PB_Unicode)
                SysFreeString_(ptrIID)
                If CurrentIID = ProgID
                  If typekind = #TKIND_COCLASS
                    For J = 0 To *oTypeAttributes\cImplTypes - 1
                      oITypeInfo\GetRefTypeOfImplType(J, @hRefType)
                      oITypeInfo\GetRefTypeInfo(hRefType, @oISubTypeInfo)
                      oISubTypeInfo\GetTypeAttr(@aSubTypeAttributes)
                      *oSubTypeAttributes = aSubTypeAttributes
                      StringFromCLSID_(*oSubTypeAttributes\guid, @ptrSubIID)
                      CurrentIID = PeekS(ptrSubIID, #PB_Any, #PB_Unicode)
                      SysFreeString_(ptrSubIID)
                      If J = 0
                        MainIID = CurrentIID
                      EndIf
                    Next
                  Else
                    MainIID = CurrentIID
                  EndIf
                EndIf
              Next
              If MainIID <> ""
                ; we have the right CLSID (i hope) and can get the pointer
                MainIID_u = StringToBStr(MainIID)
                CLSIDFromString_(MainIID_u, @CLSIDGUID)
                ptrMainIID = @CLSIDGUID
                SysFreeString_(MainIID_u)
                If Object\QueryInterface(ptrMainIID, @Object) = #S_OK
                  Err = Object\GetTypeInfo(0, 0, @ppTInfo)
                  If Err = #S_OK
                    If ppTInfo\GetTypeAttr(@pTAttr) = #S_OK
                      *oTypeAttr = pTAttr
                      If *oTypeAttr\typekind = #TKIND_DISPATCH Or *oTypeAttr\typekind = #TKIND_INTERFACE
                        If StringFromCLSID_(*oTypeAttr\guid,@IIDwStr) = #S_OK
                          SysFreeString_(IIDwStr)
                          RegCloseKey_(hKey)
                          ProcedureReturn Object
                        EndIf
                      EndIf
                    EndIf
                  EndIf
                EndIf
              EndIf
            EndIf
          EndIf
        EndIf
      RegCloseKey_(hKey)
      EndIf
    EndIf
  EndProcedure
  Macro OCX_CreateGadget(a,b,c,d,e,f, g = "")
    OCX_CreateGadget_(a,b,c,d,e,f,g)
  EndMacro 
CompilerEndIf
Procedure.l dhCreateObject2(ProgId.s, hWnd.l = 0)
  Protected object.IDispatch, Container.IUnknown
  If hWnd = 0
    dhCreateObject_(ProgId, #NUL, @object)
    ProcedureReturn object
  Else
    If AtlAXIsInit = #False
      AtlAxWinInit()
      AtlAXIsInit = #True
    EndIf
    AtlAxCreateControl(ProgId, hWnd, 0, @Container)
    If Container <> 0
      AtlAxGetControl(hWnd, @object)
      ProcedureReturn object
    EndIf
  EndIf
EndProcedure

Macro dhReleaseObject(obj)
  Define.IUnknown *object = obj
  If *object
    *object\Release()
    obj = 0
  EndIf
EndMacro
Macro dhFreeString(lpString)
  SysFreeString_(lpString)
  lpString = 0
EndMacro
Macro dhToggleExceptions(a)
  dhToggleExceptions_(a)
EndMacro
Macro dhCreateObject(a,b=0)
  dhCreateObject2(a,b)
EndMacro
Macro dhGetValue(a,b,c,d,e=0,f=0,g=0,h=0,i=0,j=0,k=0,l=0,m=0,n=0,o=0,p=0)
  dhGetValue_(a,b,c,d,e,f,g,h,i,j,k,l,m,n,o,p)
EndMacro
Macro dhPutRef(a,b,c=0,d=0,e=0,f=0,g=0,h=0,i=0,j=0,k=0,l=0,m=0,n=0)
  dhPutRef_(a,b,c,d,e,f,g,h,i,j,k,l,m,n)
EndMacro
Macro dhPutValue(a,b,c=0,d=0,e=0,f=0,g=0,h=0,i=0,j=0,k=0,l=0,m=0,n=0)
  dhPutValue_(a,b,c,d,e,f,g,h,i,j,k,l,m,n)
EndMacro
Macro dhCallMethod(a,b,c=0,d=0,e=0,f=0,g=0,h=0,i=0,j=0,k=0,l=0,m=0,n=0)
  dhCallMethod_(a,b,c,d,e,f,g,h,i,j,k,l,m,n)
EndMacro
;/////////////////////////////////////////////////////////////////////////////////
;Adjusted by srod to fix a problem with dhgetobject().
  Macro dhGetObject(a,b,c)
    dh_dhgetobjectProto(a,b,c)
  EndMacro
;/////////////////////////////////////////////////////////////////////////////////


CompilerIf Defined(DISPHELPER_NO_FOR_EACH, #PB_Variable) = #False
  Macro NEXT_(objName)
        dhReleaseObject(objName)
      Wend
    EndIf
    dhReleaseObject(xx_pEnum_xx)
  EndMacro
  Macro FOR_EACH(objName, pDisp, szMember, a=0,b=0,c=0,d=0,e=0,f=0,g=0,h=0,i=0,j=0,k=0,l=0)
    Define.IEnumVARIANT xx_pEnum_xx
    If Not dhEnumBegin(@xx_pEnum_xx, pDisp, szMember, a,b,c,d,e,f,g,h,i,j,k,l)
      While Not dhEnumNextObject(xx_pEnum_xx, @objName)
  EndMacro
CompilerEndIf

CompilerIf Defined(DISPHELPER_NO_WITH, #PB_Variable) = #False
  Macro END_WITH(objName)
    EndIf
    dhReleaseObject(objName)
  EndMacro
  Macro WITH_(objName, pDisp, szMember, a=0,b=0,c=0,d=0,e=0,f=0,g=0,h=0,i=0,j=0,k=0,l=0)
    If Not dhGetValue("%o", @objName, pDisp, szMember, a,b,c,d,e,f,g,h,i,j,k,l)
  EndMacro
CompilerEndIf

; IDE Options = PureBasic 4.20 (Windows - x86)
; CursorPosition = 326
; FirstLine = 246
; Folding = -----
; EnableXP
; HideErrorLog