; English forum: 
; Author: Richard Eikeland (updated for PB3.93 by ts-soft)
; Date: 12. April 2003
; OS: Windows
; Demo: No


; Apr. 12, 2003
; Converted to PB by Richard Eikeland
; This code is posted as is with out any waranties.
;

Import "uuid.lib"
EndImport
Import "User32.lib"
EndImport

#SERVICE_WIN32_OWN_PROCESS = $10
#SERVICE_WIN32_SHARE_PROCESS = $20
#SERVICE_WIN32 = #SERVICE_WIN32_OWN_PROCESS + #SERVICE_WIN32_SHARE_PROCESS
#SERVICE_ACCEPT_STOP = $1
#SERVICE_ACCEPT_PAUSE_CONTINUE = $2
#SERVICE_ACCEPT_SHUTDOWN = $4
#SC_MANAGER_CONNECT = $1
#SC_MANAGER_CREATE_SERVICE = $2
#SC_MANAGER_ENUMERATE_SERVICE = $4
#SC_MANAGER_LOCK = $8
#SC_MANAGER_QUERY_LOCK_STATUS = $10
#SC_MANAGER_MODIFY_BOOT_CONFIG = $20

#STANDARD_RIGHTS_REQUIRED = $F0000
#SERVICE_QUERY_CONFIG = $1
#SERVICE_CHANGE_CONFIG = $2
#SERVICE_QUERY_STATUS = $4
#SERVICE_ENUMERATE_DEPENDENTS = $8
#SERVICE_START = $10
#SERVICE_STOP = $20
#SERVICE_PAUSE_CONTINUE = $40
#SERVICE_INTERROGATE = $80
#SERVICE_USER_DEFINED_CONTROL = $100
;#SERVICE_ALL_ACCESS = #STANDARD_RIGHTS_REQUIRED | #SERVICE_QUERY_CONFIG | #SERVICE_CHANGE_CONFIG | #SERVICE_QUERY_STATUS | #SERVICE_ENUMERATE_DEPENDENTS | #SERVICE_START | #SERVICE_STOP | #SERVICE_PAUSE_CONTINUE | #SERVICE_INTERROGATE |#SERVICE_USER_DEFINED_CONTROL

#SERVICE_DEMAND_START = $3
#SERVICE_ERROR_NORMAL = $1

;- SERVICE_CONTROL
#SERVICE_CONTROL_STOP = $1
#SERVICE_CONTROL_PAUSE = $2
#SERVICE_CONTROL_CONTINUE = $3
#SERVICE_CONTROL_INTERROGATE = $4
#SERVICE_CONTROL_SHUTDOWN = $5
#SERVICE_CONTROL_POWEREVENT =$D

;-SERVICE_STATE
#SERVICE_STOPPED = $1
#SERVICE_START_PENDING = $2
#SERVICE_STOP_PENDING = $3
#SERVICE_RUNNING = $4
#SERVICE_CONTINUE_PENDING = $5
#SERVICE_PAUSE_PENDING = $6
#SERVICE_PAUSED = $7



; GUID_POWERSCHEME_PERSONALITY - 245d8541-3943-4422-b025-13A7-84F679B7
; 
; Notifies each time the active power scheme personality changes. All power schemes map To one of these personalities. the Data member is a GUID that indicates the new active power scheme personality.
; 
; GUID_MIN_POWER_SAVINGS - 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c
; 
; High Performance - the scheme is designed To deliver maximum Performance at the expense of power consumption savings.
; GUID_MAX_POWER_SAVINGS a1841308-3541-4fab-bc81-f71556f20b4a
; 
; power Saver - the scheme is designed To deliver maximum power consumption savings at the expense of system Performance And responsiveness.
; GUID_TYPICAL_POWER_SAVINGS - 381b4222-f694-41f0-9685-ff5bb260df2e
; 
; Automatic - the scheme is designed To automatically balance Performance And power consumption savings.
; 
; GUID_ACDC_POWER_SOURCE - 5d3e9a59-e9D5-4b00-a6bd-ff34ff516548
; 
; Notifies each time the system power source changes. the Data member is a DWORD With values from the SYSTEM_POWER_CONDITION Enumeration that indicates the current power source.
    ; 
    ; PoAc (zero)
    ; 
    ; the computer is powered by an AC power source (Or similar, such As a laptop powered by a 12V automotive adapter).
    ; PoDc (one)
    ; 
    ; the computer is powered by an onboard battery power source.
    ; PoHot (two)
    ; 
    ; the computer is powered by a short-term power source such As a UPS device.
    ; 
    ; GUID_BATTERY_PERCENTAGE_REMAINING - a7ad8041-b45a-4cae-87a3-eecbb468a9e1
    ; 
    ; Notifies each time the remaining battery capacity changes. the granularity varies from system To system but the finest granularity is 1 percent. the Data member is a DWORD that indicates the current battery capacity remaining, in percentage from 0 through 100. the Data member has no information And can be ignored.
    ; GUID_IDLE_BACKGROUND_TASK - 515c31d8-f734-163d-a0fd-11a0-8c91e8f1
    ; 
    ; the idle/background task notification is delivered when the system is busy. This indicates that the system will Not be moving into an idle state in the near future And that the current time is a good time For components To perform background Or idle tasks that would otherwise prevent the computer from entering an idle state. There is no notification when the system is able To move into an idle state. the background idle task notification does Not indicate If a user is present Or Not the computer. the Data member has no information And can be ignored.
        ; GUID_SYSTEM_AWAYMODE - 98a7f580-01f7-48aa-9c0f-44352c29e5C0
        ; 98a7f58001f748aa9c0f44352c29e5C0
        ; the away mode notification indicates that the system is entering Or exiting away mode. the Data member is a DWORD that indicates the current away mode state.
        ; 
        ; 0x0
        ; 
        ; the computer is exiting away mode.
        ; 0x1
        ; 
        ; the computer is entering away mode.
        ; 
        ; GUID_MONITOR_POWER_ON - 02731015-4510-4526-99e6-e5a17ebd1aea
        ; 
        ; the monitor on/off notification indicates when the primary system monitor is on Or off. This notification is useful For components that actively render content To the display device, such As media visualization. these applications should register For This notification And stop rendering graphics content when the monitor is off To reduce system power consumption. the Data member is a DWORD that indicates the current monitor state.
            ; 
            ; 0x0
            ; 
            ; the monitor is off.
            ; 0x1
            ; 
            ; the monitor is on.

;#GUID_SYSTEM_AWAYMODE = $98A7F58001F748AA9C0F44352C29E5C0
DataSection
  
  GUID_SYSTEM_AWAYMODE: ; {98a7f580-01f7-48aa-9c0f-44352c29e5C0}
  Data.l $98A7F580
  Data.w $01F7, $48AA
  Data.b $9C, $0F, $44, $35, $2C, $29, $E5, $C0
  
EndDataSection
; Structure GUID
  ; Data1.l
  ; Data2.w
  ; Data3.w
  ; Data4.c[8]
;EndStructure

Structure POWERBROADCAST_SETTING
  PowerSetting.GUID
  DataLength.l
  Data.c
EndStructure

Global ServiceStatus.SERVICE_STATUS, hServiceStatus.l, hPowerNotify.l, SERVICE_NAME.s, Finish.l, PBS.POWERBROADCAST_SETTING
Declare Handler(fdwControl.l)
Declare HandlerEx(dwControl.l, dwEventType.l, lpEventData.l, lpContext.l)
Declare ServiceMain(dwArgc.l, lpszArgv.l)
Declare WriteLog(Value.s)

Procedure Main()
  
  hSCManager.l
  hService.l
  ServiceTableEntry.SERVICE_TABLE_ENTRY
  b.l
  cmd.s
  ;Change SERVICE_NAME and app name as needed
  AppPath.s = "C:\Dokumente und Einstellungen\injk\Eigene Dateien\PMEvent\NTService.exe"
  SERVICE_NAME = "NTService"
  
  cmd = Trim(LCase(ProgramParameter()))
  Select cmd
  Case "install" ;Install service on machine
    hSCManager = OpenSCManager_(0, 0, #SC_MANAGER_CREATE_SERVICE)
    hService = CreateService_(hSCManager, SERVICE_NAME, SERVICE_NAME, #SERVICE_ALL_ACCESS, #SERVICE_WIN32_OWN_PROCESS, #SERVICE_DEMAND_START, #SERVICE_ERROR_NORMAL, AppPath, 0, 0, 0, 0, 0)
    CloseServiceHandle_(hService)
    CloseServiceHandle_(hSCManager)
    Finish = 1
  Case "uninstall" ;Remove service from machine
    hSCManager = OpenSCManager_(0, 0, #SC_MANAGER_CREATE_SERVICE)
    hService = OpenService_(hSCManager, SERVICE_NAME, #SERVICE_ALL_ACCESS)
    DeleteService_(hService)
    CloseServiceHandle_(hService)
    CloseServiceHandle_(hSCManager)
    Finish = 1
  Default
    *sname.s = SERVICE_NAME
    ;Start the service
    ServiceTableEntry\lpServiceName = @SERVICE_NAME
    ServiceTableEntry\lpServiceProc = @ServiceMain()
    b = StartServiceCtrlDispatcher_(@ServiceTableEntry)
    WriteLog("Starting Service bResult=" + Str(b))
    If b = 0
      Finish = 1
    EndIf
  EndSelect
  
  Repeat
    
  Until Finish =1
  
  End
  
EndProcedure

Procedure Handler(fdwControl.l)
  b.l
  
  Select fdwControl
  Case #SERVICE_CONTROL_PAUSE
    ;** Do whatever it takes To pause here.
    ServiceStatus\dwCurrentState = #SERVICE_PAUSED
  Case #SERVICE_CONTROL_CONTINUE
    ;** Do whatever it takes To continue here.
    ServiceStatus\dwCurrentState = #SERVICE_RUNNING
  Case #SERVICE_CONTROL_STOP
    ServiceStatus\dwWin32ExitCode = 0
    ServiceStatus\dwCurrentState = #SERVICE_STOP_PENDING
    ServiceStatus\dwCheckPoint = 0
    ServiceStatus\dwWaitHint = 0 ;Might want a time estimate
    b = SetServiceStatus_(hServiceStatus, ServiceStatus)
    ;** Do whatever it takes to stop here.
    Finish = 1
    ServiceStatus\dwCurrentState = #SERVICE_STOPPED
  Case #SERVICE_CONTROL_INTERROGATE
    ;Fall through To send current status.
    Finish = 1
    ;Else
  Case #SERVICE_CONTROL_POWEREVENT
  EndSelect
  ;Send current status.
  b = SetServiceStatus_(hServiceStatus, ServiceStatus)
EndProcedure

Procedure HandlerEx(dwControl.l, dwEventType.l, lpEventData.l, lpContext.l)
  b.l
  
  Select dwControl
    Case #SERVICE_CONTROL_PAUSE
      ;** Do whatever it takes To pause here.
      ServiceStatus\dwCurrentState = #SERVICE_PAUSED
    Case #SERVICE_CONTROL_CONTINUE
      ;** Do whatever it takes To continue here.
      ServiceStatus\dwCurrentState = #SERVICE_RUNNING
    Case #SERVICE_CONTROL_STOP
      ServiceStatus\dwWin32ExitCode = 0
      ServiceStatus\dwCurrentState = #SERVICE_STOP_PENDING
      ServiceStatus\dwCheckPoint = 0
      ServiceStatus\dwWaitHint = 0 ;Might want a time estimate
      b = SetServiceStatus_(hServiceStatus, ServiceStatus)
      ;** Do whatever it takes to stop here.
      If UnregisterPowerSettingNotification(hPowerNotify)=0
        WriteLog("UnregisterPowerNotofication failed.")
      EndIf
      Finish = 1
      ServiceStatus\dwCurrentState = #SERVICE_STOPPED
    Case #SERVICE_CONTROL_INTERROGATE
      ;Fall through To send current status.
      Finish = 1
      ;Else
    Case #SERVICE_CONTROL_POWEREVENT
      WriteLog("SERVICE_CONTROL_POWEREVENT event occured")
  EndSelect
  ;Send current status.
  b = SetServiceStatus_(hServiceStatus, ServiceStatus)
EndProcedure

Procedure ServiceMain(dwArgc.l, lpszArgv.l)
  b.l
  WriteLog("ServiceMain")
  ;Set initial state
  ServiceStatus\dwServiceType = #SERVICE_WIN32_OWN_PROCESS
  ServiceStatus\dwCurrentState = #SERVICE_START_PENDING
  ServiceStatus\dwControlsAccepted = #SERVICE_ACCEPT_STOP | #SERVICE_ACCEPT_PAUSE_CONTINUE | #SERVICE_ACCEPT_SHUTDOWN | #SERVICE_ACCEPT_POWEREVENT
  ServiceStatus\dwWin32ExitCode = 0
  ServiceStatus\dwServiceSpecificExitCode = 0
  ServiceStatus\dwCheckPoint = 0
  ServiceStatus\dwWaitHint = 0
  
  hServiceStatus = RegisterServiceCtrlHandler_(SERVICE_NAME, @HandlerEx())
  ServiceStatus\dwCurrentState = #SERVICE_START_PENDING
  b = SetServiceStatus_(hServiceStatus, ServiceStatus)
  
  ;** Do Initialization Here
  hPowerNotify = RegisterPowerSettings_(hServiceStatus, @GUID_SYSTEM_AWAYMODE, 1) ; 1=DEVICE_NOTIFY_SERVICE_HANDLE
  
  ServiceStatus\dwCurrentState = #SERVICE_RUNNING
  b = SetServiceStatus_(hServiceStatus, ServiceStatus)
  
  ;** Perform tasks -- If none exit
  ;hwnd = GetDesktopWindow_()
  
  ; If OpenWindow(0, 200, 300, 400, 200,"QuietHdSVC",#PB_Window_SystemMenu | #PB_Window_MinimizeGadget | #PB_Window_MaximizeGadget) ; | #PB_Window_Invisible)
    ; WriteLog("Window opened")
    ; tray = AddSysTrayIcon(0, WindowID(0), LoadImage(0, "quiethd.ico"))
    ; WriteLog("Tray added: "+Str(tray))
    ; SysTrayIconToolTip (tray, "QuietHD")
    ; Repeat
      ; 
      ; EventID.l=WaitWindowEvent()
      ; If EventID=#PB_Event_CloseWindow   
        ; quit=#True
        ; 
      ; EndIf
    ; Until quit=#True
  ; EndIf
  ;** If an error occurs the following should be used for shutting
  ;** down:
  ; SetServerStatus SERVICE_STOP_PENDING
  ; Clean up
  ; SetServerStatus SERVICE_STOPPED
EndProcedure

Procedure WriteLog(Value.s)
  sfile.s = SERVICE_NAME+"Log.txt"
  If FileSize(sfile)>0
    File = OpenFile(#PB_Any, sfile)
  Else
    File = CreateFile(#PB_Any, sfile)
  EndIf
  If File
    FileSeek(File, Lof(File)
    WriteStringN(File, FormatDate("%YYYY-%MM-%DD %HH:%II:%SS - ", Date())+Value)
    CloseFile(File)
  EndIf
EndProcedure

Main() 
; jaPBe Version=3.8.9.728
; FoldLines=00C200DD
; Build=15
; FirstLine=158
; CursorPosition=148
; ExecutableFormat=Windows
; Executable=C:\Dokumente und Einstellungen\injk\Eigene Dateien\PMEvent\NTService.exe
; DontSaveDeclare
; EOF