; ---------------------------------------------------------------------
;
; Receive COM events from ActiveX controls
;
; 26.4.07 by Timo 'fr34k' Harter
;
; ---------------------------------------------------------------------
;
; Connection = OCX_ConnectEvents(Object, @EventCallback())
;
;   Connects the callback to the Object. Returns an ID if successful
;   that can be used to disconnect the callback again.
;   If the callback is not disconnected before the object is distroyed, this is
;   done automatically.
;
;   Returns 0 if the connection failed.
;
;   NOTE: The Object must expose a ConnectionPoint, and it must provide
;         information about it through a type library!
;
;
; OCX_DisconnectEvents(Object, Connection)
;
;   Disconnects a previous set event callback from the object.
;   If the callback is not disconnected before the object is distroyed, this is
;   done automatically.
;
; The Callback has to look like this:
;
;   Procedure EventCallback(Event$, ParameterCount, *Parameters)
;   EndProcedure
;
;   Event$         - the name of the event
;   ParameterCount - the number of parameters for the event
;   *Parameters    - the parameter list. the functions below can be used to access this
;
; OCX_EventLong(*Parameters, index)
; OCX_EventFloat(*Parameters, index)
; OCX_EventQuad(*Parameters, index)
; OCX_EventDouble(*Parameters, index)
; OCX_EventString(*Parameters, index)
; OCX_EventObject(*Parameters, index)
;
;   Returns the parameter at position 'index' (one-based) in form of a
;   long, float, quad, double, string or object.
;   These functions are only valid inside the callback with the *Parameters parameter.
;
;   For OCX_EventObject(), the returned object must be released!
;
; ---------------------------------------------------------------------

EnableExplicit

Prototype EventCallback(EventName$, ParameterCount, *Parameters)

Structure EventSink
  *Vtbl.l
  RefCount.l 
  ConnIID.IID
  TypeInfo.ITypeInfo
  Callback.EventCallback
EndStructure

Procedure EventSink_QueryInterface(*THIS.EventSink, *IID.IID, *Object.LONG) 
  If CompareMemory(*IID, ?IID_IUnknown, SizeOf(IID)) Or CompareMemory(*IID, ?IID_IDispatch, SizeOf(IID)) Or CompareMemory(*IID, @*THIS\ConnIID, SizeOf(IID))
    *Object\l = *THIS
    *THIS\RefCount + 1
    ProcedureReturn #S_OK
  Else
    *Object\l = 0
    ProcedureReturn #E_NOINTERFACE
  EndIf
EndProcedure

Procedure EventSink_AddRef(*THIS.EventSink)
  *THIS\RefCount + 1
  ProcedureReturn *THIS\RefCount
EndProcedure

Procedure EventSink_Release(*THIS.EventSink)
  *THIS\RefCount - 1
 
  If *THIS\RefCount = 0
    *THIS\TypeInfo\Release()
    FreeMemory(*THIS)
    ProcedureReturn 0
  Else 
    ProcedureReturn *THIS\RefCount
  EndIf
EndProcedure

Procedure EventSink_GetTypeInfoCount(*THIS.EventSink, *pctinfo.LONG)
  *pctinfo\l = 1
  ProcedureReturn #S_OK
EndProcedure

Procedure EventSink_GetTypeInfo(*THIS.EventSink, iTInfo, lcid, *ppTInfo.LONG)
  *ppTInfo\l = *THIS\TypeInfo
  *THIS\TypeInfo\AddRef()
  ProcedureReturn #S_OK
EndProcedure

Procedure EventSink_GetIDsOfNames(*THIS.EventSink, *riid, *rgszNames, *cNames, lcid, *DispID)
  ProcedureReturn DispGetIDsOfNames_(*THIS\TypeInfo, *rgszNames, *cNames, *DispID)
EndProcedure

Procedure EventSink_Invoke(*THIS.EventSink, dispid, *riid, lcid, wflags.w, *Params.DISPPARAMS, *Result.VARIANT, *pExept, *ArgErr)
  Protected NameCount, bstrName
  Protected Callback.EventCallback = *THIS\Callback ; work around a compiler bug
 
  If Callback And *THIS\TypeInfo\GetNames(dispid, @bstrName, 1, @NameCount) = #S_OK And bstrName
    Callback(PeekS(bstrName, -1, #PB_Unicode), *Params\cArgs + *Params\cNamedArgs, *Params)
    SysFreeString_(bstrName)   
  EndIf
 
  ProcedureReturn #S_OK
EndProcedure

; ---------------------------------------------------------------------

Procedure.l OCX_EventLong(*Params.DISPPARAMS, index)
  Protected Value.VARIANT, puArgErr
 
  If index > 0 And index <= *Params\cArgs+*Params\cNamedArgs
    DispGetParam_(*Params, index-1, #VT_I4, @Value, @puArgErr)
  EndIf
 
  ProcedureReturn Value\lVal
EndProcedure

Procedure.f OCX_EventFloat(*Params.DISPPARAMS, index)
  Protected Value.VARIANT, puArgErr
 
  If index > 0 And index <= *Params\cArgs+*Params\cNamedArgs
    DispGetParam_(*Params, index-1, #VT_R4, @Value, @puArgErr)
  EndIf
 
  ProcedureReturn Value\fltVal
EndProcedure

Procedure.q OCX_EventQuad(*Params.DISPPARAMS, index)
  Protected Value.VARIANT, puArgErr
 
  If index > 0 And index <= *Params\cArgs+*Params\cNamedArgs
    DispGetParam_(*Params, index-1, #VT_I8, @Value, @puArgErr)
  EndIf
 
  ProcedureReturn Value\llVal
EndProcedure

Procedure.d OCX_EventDouble(*Params.DISPPARAMS, index)
  Protected Value.VARIANT, puArgErr
 
  If index > 0 And index <= *Params\cArgs+*Params\cNamedArgs
    DispGetParam_(*Params, index-1, #VT_R8, @Value, @puArgErr)
  EndIf
 
  ProcedureReturn Value\dblVal
EndProcedure

Procedure.s OCX_EventString(*Params.DISPPARAMS, index)
  Protected Value.VARIANT, puArgErr
  Protected Result$ = ""
 
  If index > 0 And index <= *Params\cArgs+*Params\cNamedArgs
    If DispGetParam_(*Params, index-1, #VT_BSTR, @Value, @puArgErr) = #S_OK
      If Value\bstrVal
        Result$ = PeekS(Value\bstrVal, -1, #PB_Unicode)
        SysFreeString_(Value\bstrVal)
      EndIf
    EndIf
  EndIf
 
  ProcedureReturn Result$
EndProcedure

Procedure.l OCX_EventObject(*Params.DISPPARAMS, index)
  Protected Value.VARIANT, puArgErr
 
  If index > 0 And index <= *Params\cArgs+*Params\cNamedArgs
    DispGetParam_(*Params, index-1, #VT_DISPATCH, @Value, @puArgErr)
  EndIf
 
  ProcedureReturn Value\pdispVal
EndProcedure

; ---------------------------------------------------------------------

Procedure OCX_ConnectEvents(Object.IUnknown, Callback.EventCallback) 
  Protected Container.IConnectionPointContainer
  Protected Connection.IConnectionPoint
  Protected Enum.IEnumConnectionPoints
  Protected Dispatch.IDispatch, TypeLib.ITypeLib
  Protected DispTypeInfo.ITypeInfo, TypeInfo.ITypeInfo
  Protected ConnIID.IID, *NewSink.EventSink, NewSink.IDispatch
  Protected InfoCount = 0, index, Result = 0

  If Object\QueryInterface(?IID_IConnectionPointContainer, @Container.IConnectionPointContainer) = #S_OK
    If Container\EnumConnectionPoints(@Enum.IEnumConnectionPoints) = #S_OK
      If Enum\Reset() = #S_OK And Enum\Next(1, @Connection.IConnectionPoint, #Null) = #S_OK     
        If Connection\GetConnectionInterface(@ConnIID.IID) = #S_OK
       
          If Object\QueryInterface(?IID_IDispatch, @Dispatch.IDispatch) = #S_OK         
            If Dispatch\GetTypeInfoCount(@InfoCount) = #S_OK And InfoCount = 1
              If Dispatch\GetTypeInfo(0, 0, @DispTypeInfo.ITypeInfo) = #S_OK
                If DispTypeInfo\GetContainingTypeLib(@TypeLib.ITypeLib, @index) = #S_OK
                  If TypeLib\GetTypeInfoOfGuid(@ConnIID, @TypeInfo.ITypeInfo) = #S_OK
                   
                    *NewSink.EventSink = AllocateMemory(SizeOf(EventSink))                     
                    If *NewSink
                      *NewSink\Vtbl     = ?EventSink_Vtbl
                      *NewSink\RefCount = 1
                      *NewSink\TypeInfo = TypeInfo
                      *NewSink\Callback = Callback
                      *NewSink\TypeInfo\AddRef()
                      CopyMemory(@ConnIID, @*NewSink\ConnIID, SizeOf(IID))
                      NewSink.IDispatch = *NewSink
                   
                      Connection\Advise(NewSink, @Result)
                      NewSink\Release()
                    EndIf

                    TypeInfo\Release()
                  EndIf               
                  TypeLib\Release()
                EndIf               
                DispTypeInfo\Release()
              EndIf
            EndIf                   
            Dispatch\Release()
          EndIf
       
        EndIf       
        Connection\Release()
      EndIf
      Enum\Release()
    EndIf
    Container\Release()
  EndIf
 
  ProcedureReturn Result
EndProcedure

Procedure OCX_DisconnectEvents(Object.IUnknown, EventConnection)
  Protected Container.IConnectionPointContainer
  Protected Connection.IConnectionPoint
  Protected Enum.IEnumConnectionPoints
 
  If Object\QueryInterface(?IID_IConnectionPointContainer, @Container.IConnectionPointContainer) = #S_OK
    If Container\EnumConnectionPoints(@Enum.IEnumConnectionPoints) = #S_OK
      If Enum\Reset() = #S_OK And Enum\Next(1, @Connection.IConnectionPoint, #Null) = #S_OK     
        Connection\Unadvise(EventConnection)
        Connection\Release()
      EndIf
      Enum\Release()
    EndIf
    Container\Release()
  EndIf       
EndProcedure

; ---------------------------------------------------------------------

DataSection
  EventSink_Vtbl:
    Data.l @EventSink_QueryInterface()
    Data.l @EventSink_AddRef()
    Data.l @EventSink_Release()
    Data.l @EventSink_GetTypeInfoCount()
    Data.l @EventSink_GetTypeInfo()
    Data.l @EventSink_GetIDsOfNames()
    Data.l @EventSink_Invoke()   

  IID_IUnknown: ; {00000000-0000-0000-C000-000000000046}
    Data.l $00000000
    Data.w $0000, $0000
    Data.b $C0, $00, $00, $00, $00, $00, $00, $46

  IID_IDispatch: ; {00020400-0000-0000-C000-000000000046}
    Data.l $00020400
    Data.w $0000, $0000
    Data.b $C0, $00, $00, $00, $00, $00, $00, $46   
   
  IID_IConnectionPointContainer: ; {B196B284-BAB4-101A-B69C-00AA00341D07}
    Data.l $B196B284
    Data.w $BAB4, $101A
    Data.b $B6, $9C, $00, $AA, $00, $34, $1D, $07
   
EndDataSection 
; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 287
; FirstLine = 241
; Folding = ---
; EnableXP
; HideErrorLog