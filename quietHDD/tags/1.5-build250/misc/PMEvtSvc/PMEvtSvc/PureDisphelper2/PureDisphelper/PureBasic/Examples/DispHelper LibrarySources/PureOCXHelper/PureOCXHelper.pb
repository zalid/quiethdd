Import "atl.lib"
  AtlAxCreateControl(lpszName.l,hWnd.l,*pStream.IStream,*ppUnkContainer.IUnknown) As "_AtlAxCreateControl"
  AtlAxGetControl(hWnd.l,*pp.IUnknown) As "_AtlAxGetControl"
  AtlAxWinInit() As "_AtlAxWinInit"
EndImport

EnableExplicit

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

ProcedureDLL PureOCXHelper_Init()
  CoInitialize_(0)
  AtlAxWinInit()
EndProcedure

ProcedureDLL PureOCXHelper_End()
  CoUninitialize_()
EndProcedure

ProcedureDLL OCX_CreateGadget2(Gadget.l, x.l, y.l, width.l, height.l, ProgID.s, CLSID.s)
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
    Gadget = ContainerGadget(#PB_Any, x, y, width, 	height)
  Else
    ContainerGadget(Gadget, x, y, width, 	height)
  EndIf
  CloseGadgetList()
  
  ProgID_u = StringToBStr(ProgID)
  AtlAxCreateControl(ProgID_u, GadgetID(Gadget), 0, @Container)
  SysFreeString_(ProgID_u)
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

ProcedureDLL OCX_CreateGadget(Gadget.l, x.l, y.l, width.l, height.l, ProgID.s)
  ProcedureReturn OCX_CreateGadget2(Gadget, x, y, width, height, ProgID, "")
EndProcedure


; IDE Options = PureBasic 4.20 Beta 2 (Windows - x86)
; CursorPosition = 91
; FirstLine = 44
; Folding = --
; HideErrorLog