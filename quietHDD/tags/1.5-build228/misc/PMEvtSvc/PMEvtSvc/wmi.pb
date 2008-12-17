;- WMI Intialisierung, Datenabruf und Deinitialisierung
; rewritten for PB4 and use as an "IncludeFile"
; save this code as "wmi.pb"
; use it in your project with:
;
; includefile "wmi.pb"
; WMI_INIT()
; WMI_Call("Select * FROM Win32_OperatingSystem", "Caption, CSDVersion, SerialNumber, RegisteredUser, Organization")
; ResetList(wmidata())
; While NextElement(wmidata())
;   Debug wmidata()  ; Alle Listenelemente darstellen / show all elements
; Wend
; WMI_RELEASE("OK")
;
;----------------------------------------------------------------------------------------------------------------
; Update für PB4 Final by ts-soft
; unnötige Konstanten und Structuren entfernt (sind in PB enthalten)
; voller Unicode support
; ---------------------------------------------------------------------------------------------------------------
;- KONSTANTEN  PROZEDUREN  STRUKTUREN

#COINIT_MULTITHREAD = 0
#RPC_C_AUTHN_LEVEL_CONNECT = 2
#RPC_C_IMP_LEVEL_IDENTIFY = 2
#EOAC_NONE = 0
#RPC_C_AUTHN_WINNT = 10
#RPC_C_AUTHZ_NONE = 0
#RPC_C_AUTHN_LEVEL_CALL = 3
#RPC_C_IMP_LEVEL_IMPERSONATE = 3
#CLSCTX_INPROC_SERVER = 1
#wbemFlagReturnImmediately = 16
#wbemFlagForwardOnly = 32
#IFlags = #wbemFlagReturnImmediately + #wbemFlagForwardOnly
#WBEM_INFINITE = $FFFFFFFF
#WMISeparator = ","

Procedure StringToBStr(string$)
  Protected Unicode$ = Space(StringByteLength(string$, #PB_Unicode) + 1)
  Protected bstr_string.l
  PokeS(@Unicode$, String$, -1, #PB_Unicode)
  bstr_string = SysAllocString_(@Unicode$)
  ProcedureReturn bstr_string
EndProcedure

Procedure.s UniToPB(*Unicode)
  ProcedureReturn PeekS(*Unicode, #PB_Any, #PB_Unicode)
EndProcedure


Global txt$, loc.IWbemLocator, svc.IWbemServices, svc.NextEvent, pEnumerator.IEnumWbemClassObject, pclsObj.IWbemClassObject, x.Variant, error.l
Global NewList wmidata.s()

ProcedureDLL.s wmi_release(dumdum$)
  ;- WMI Release
  svc\release()
  loc\release()
  pEnumerator\release()
  If error=0
    pclsObj\release()
  EndIf
  CoUninitialize_()
  If FindString(dumdum$, "ERROR", 1)
    MessageRequester("", dumdum$)
    End
  EndIf
EndProcedure

ProcedureDLL.s wmi_init()
  ;- WMI Initialize
  txt$=""
  CoInitializeEx_(0,#COINIT_MULTITHREAD)
  hres=CoInitializeSecurity_(0, -1,0,0,#RPC_C_AUTHN_LEVEL_CONNECT,#RPC_C_IMP_LEVEL_IDENTIFY,0,#EOAC_NONE,0)
  If hres <> 0: txt$="ERROR: unable To call CoInitializeSecurity": wmi_release(txt$): EndIf
  hres=CoCreateInstance_(?CLSID_WbemLocator,0,#CLSCTX_INPROC_SERVER,?IID_IWbemLocator,@loc.IWbemLocator)
  If hres <> 0: txt$="ERROR: unable To call CoCreateInstance": wmi_release(txt$): EndIf
  hres=loc\ConnectServer(StringToBStr("root\cimv2"),0,0,0,0,0,0,@svc.IWbemServices)
  If hres <> 0: txt$="ERROR: unable To call IWbemLocator::ConnectServer": wmi_release(txt$): EndIf
  hres=svc\queryinterface(?IID_IUnknown,@pUnk.IUnknown)
  hres=CoSetProxyBlanket_(svc,#RPC_C_AUTHN_WINNT,#RPC_C_AUTHZ_NONE,0,#RPC_C_AUTHN_LEVEL_CALL,#RPC_C_IMP_LEVEL_IMPERSONATE,0,#EOAC_NONE)
  If hres <> 0: txt$="ERROR: unable To call CoSetProxyBlanket": wmi_release(txt$): EndIf
  hres=CoSetProxyBlanket_(pUnk,#RPC_C_AUTHN_WINNT,#RPC_C_AUTHZ_NONE,0,#RPC_C_AUTHN_LEVEL_CALL,#RPC_C_IMP_LEVEL_IMPERSONATE,0,#EOAC_NONE)
  If hres <> 0: txt$="ERROR: unable To call CoSetProxyBlanket": wmi_release(txt$): EndIf
  pUnk\release()
  ProcedureReturn txt$
EndProcedure

ProcedureDLL.s WMI_Call(WMISelect.s, WMICommand.s)
  ;- Call Data
  OnErrorResume()
  error=0
  WMICommand=WMISelect+","+WMICommand
  ClearList(wmidata())
  k=CountString(WMICommand,#WMISeparator)
  Dim wmitxt$(k)
  For i=0 To k
    wmitxt$(i) = Trim(StringField(WMICommand,i+1,#WMISeparator))
  Next

  ;hres=svc\ExecQuery(StringToBStr("WQL"),StringToBStr(wmitxt$(0)), #IFlags,0,@pEnumerator.IEnumWbemClassObject)
  hres=svc\ExecNotificationQuery(StringToBStr("WQL"),StringToBStr(wmitxt$(0)), #IFlags,0,@pEnumerator.IEnumWbemClassObject)
  If hres <> 0: txt$="ERROR: unable To call IWbemServices::ExecQuery": wmi_release(txt$): EndIf
  hres=svc\NextEvent()
  Debug hres
  hres=svc\NextEvent()
  Debug hres
  hres=svc\NextEvent()
  Debug hres
  hres=svc\NextEvent()
  Debug hres
  
  hres=pEnumerator\reset()
  Repeat
    hres=pEnumerator\Next(#WBEM_INFINITE, 1, @pclsObj.IWbemClassObject, @uReturn)
    For i=1 To k
      Sleep_(0)
      If uReturn <> 0

        hres=pclsObj\get(StringToBStr(wmitxt$(i)), 0, @x.Variant, 0, 0)

        type=x\vt

        Select type

          Case 8200
            val.s=""
            nDim=SafeArrayGetDim_(x\lVal)
            SafeArrayGetUBound_(x\lVal, nDim, @plUbound)
            For z=0 To plUbound
              SafeArrayGetElement_(x\lVal, @z, @rgVar)
              val.s=val.s+", "+UniToPB(rgVar)
            Next
            val.s=Mid(val.s, 3, Len(val.s))

          Case 8195
            val.s=""
            nDim=SafeArrayGetDim_(x\scode)
            SafeArrayGetUBound_(x\scode, nDim, @plUbound)
            For z=0 To plUbound
              SafeArrayGetElement_(x\scode, @z, @rgVar)
              val.s=val.s + ", " +  Str(rgVar)
            Next
            val.s=Mid(val.s, 3, Len(val.s))

          Case 11
            If x\boolVal=0
              val.s="FALSE"
            ElseIf x\boolVal=-1
              val.s="TRUE"
            EndIf

          Case 8
            val.s=UniToPB(x\bstrVal)

          Case 3
            val.s=Str(x\lVal)

          Case 1
            val.s="n/a"

          Default
            val.s=""

        EndSelect

        If FindString(wmitxt$(i), "Date", 1) Or FindString(wmitxt$(i), "Time", 1)
          AddElement(wmidata())
          wmidata()=Mid(val, 7, 2)+"."+Mid(val, 5, 2)+"."+Mid(val, 1, 4)+" "+Mid(val, 9, 2)+":"+Mid(val, 11,2)+":"+Mid(val, 13,2) ;+Chr(10)+Chr(13)
        Else
          AddElement(wmidata())
          wmidata()=Trim(val) ;+Chr(10)+Chr(13)
        EndIf
      EndIf
    Next

  Until uReturn = 0
  If CountList(wmidata())=0
    For i=1 To k
      AddElement(wmidata())
      wmidata()="n/a"
    Next
    error=1
  EndIf
  ProcedureReturn wmidata()
EndProcedure

;- Data Section
DataSection
  CLSID_IEnumWbemClassObject:
  ;1B1CAD8C-2DAB-11D2-B604-00104B703EFD
  Data.l $1B1CAD8C
  Data.w $2DAB, $11D2
  Data.b $B6, $04, $00, $10, $4B, $70, $3E, $FD
  IID_IEnumWbemClassObject:
  ;7C857801-7381-11CF-884D-00AA004B2E24
  Data.l $7C857801
  Data.w $7381, $11CF
  Data.b $88, $4D, $00, $AA, $00, $4B, $2E, $24
  CLSID_WbemLocator:
  ;4590f811-1d3a-11d0-891f-00aa004b2e24
  Data.l $4590F811
  Data.w $1D3A, $11D0
  Data.b $89, $1F, $00, $AA, $00, $4B, $2E, $24
  IID_IWbemLocator:
  ;dc12a687-737f-11cf-884d-00aa004b2e24
  Data.l $DC12A687
  Data.w $737F, $11CF
  Data.b $88, $4D, $00, $AA, $00, $4B, $2E, $24
  IID_IUnknown:
  ;00000000-0000-0000-C000-000000000046
  Data.l $00000000
  Data.w $0000, $0000
  Data.b $C0, $00, $00, $00, $00, $00, $00, $46

EndDataSection

; IDE Options = PureBasic 4.20 (Windows - x86)
; CursorPosition = 49
; FirstLine = 39
; Folding = -
; EnableAsm
; EnableUnicode
; EnableThread
; EnableOnError
; Executable = PMEvent.exe