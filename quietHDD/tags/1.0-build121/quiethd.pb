
#BROADCAST_QUERY_DENY       = $424D5144
#PBT_APMQUERYSUSPEND        = $00
#WM_POWERBROADCAST          = $218
#PBT_APMBATTERYLOW          = $09
#PBT_APMPOWERSTATUSCHANGE   = $0A
#PBT_APMOEMEVENT            = $0B
#PBT_APMQUERYSUSPEND        = $00
#PBT_APMQUERYSUSPENDFAILED  = $02
#PBT_APMRESUMECRITICAL      = $06
#PBT_APMRESUMEAUTOMATIC     = $12
#PBT_APMQUERYSUSPENDFAILED  = $02

#IOCTL_IDE_PASS_THROUGH     = $04D028
#IOCTL_ATA_PASS_THROUGH     = $04D02C
#HDIO_DRIVE_CMD             = $03F1   ; Taken from /usr/include/linux/hdreg.h
#WIN_SETFEATURES            = $EF
#SETFEATURES_EN_AAM         = $42
#SETFEATURES_DIS_AAM        = $C2
#SETFEATURES_EN_APM         = $05
#SETFEATURES_DIS_APM        = $85

#ATA_FLAGS_DRDY_REQUIRED    = $01
#ATA_FLAGS_DATA_IN          = $02
#ATA_FLAGS_DATA_OUT         = $04
#ATA_FLAGS_48BIT_COMMAND    = $08


Macro ZHex(val,digits=2)
  RSet(Hex(val), digits, "0")
EndMacro
Macro CR
  Chr(10)
EndMacro

Enumeration
  #inFeaturesReg      ;0
  #inSectorCountReg   ;1 
  #inSectorNumberReg  ;2
  #inCylLowReg        ;3
  #inCylHighReg       ;4
  #inDriveHeadReg     ;5          
  #inCommandReg       ;6
  #inReserved         ;7
EndEnumeration

Enumeration
  #outErrorReg         ;0
  #outSectorCountReg   ;1 
  #outSectorNumberReg  ;2
  #outCylLowReg        ;3
  #outCylHighReg       ;4
  #outDriveHeadReg     ;5          
  #outStatusReg        ;6
  #outReserved         ;7
EndEnumeration

Structure IDEREGS   ; this is also ATA_PASS_THROUGH?! IDEREGS is without DataBuffersize and DataBuffer
  bFeaturesReg.c      ;0
  bSectorCountReg.c   ;1 
  bSectorNumberReg.c  ;2
  bCylLowReg.c        ;3
  bCylHighReg.c       ;4
  bDriveHeadReg.c     ;5          
  bCommandReg.c       ;6
  bReserved.c         ;7
  DataBufferSize.l    ;8
  DataBuffer.c[1]     ;9, 10
EndStructure

Structure ATA_PASS_THROUGH_EX
  length.w
  AtaFlags.w
  PathId.c
  TargetId.c
  Lun.c
  ReservedAsUchar.c
  DataTransferLength.l
  TimeOutValue.l
  ReservedAsUlong.l
  DataBufferOffset.l
  PreviousTaskFile.c[8]
  CurrentTaskFile.c[8]
EndStructure

Structure ATA_PASS_THROUGH_EX_WITH_BUFFERS
  length.w
  AtaFlags.w
  PathId.c
  TargetId.c
  Lun.c
  ReservedAsUchar.c
  DataTransferLength.l
  TimeOutValue.l
  ReservedAsUlong.l
  DataBufferOffset.l
  PreviousTaskFile.c[8]
  CurrentTaskFile.c[8]
  Filler.l
  ucDataBuf.c[512]
EndStructure

Global DisableSuspend = 0
Global pdebug=#False, tray=#True, help.s, APMValue.l=255, warnuser=#True

help.s =          "eeeHDD Build:" + Str(#jaPBe_ExecuteBuild)+CR+CR
help.s = help.s + "Homepage http://sites.google.com/site/eeehddsite/"+CR+CR
help.s = help.s + "eeeHDD disables/modifies the harddrives APM feature setting"+CR
help.s = help.s + "Modifying APM value also controls HDD's spindown rate. Do not let your HDD"+CR
help.s = help.s + "spindown/spinup too often when you set small values (less 128) as it"+CR
help.s = help.s + "shortens then life of your HDD!  You have been warned!"+CR+CR
help.s = help.s + "Usage:"+CR
help.s = help.s + "eeehdd.exe [/DEBUG] [/NOTRAY] [/APMVALUE:n] [/NOWARN]"+CR
help.s = help.s + "  /DEBUG       Opens a console window for debug output"+CR
help.s = help.s + "  /NOTRAY      Do not add systray icon. In this case eeeHDD"+CR
help.s = help.s + "               can be terminated in the taskmanager only."+CR
help.s = help.s + "  /APMVALUE:n  Do not disable APM but set APM it to this value."+CR
help.s = help.s + "               Valid values are in range between 1-255"+CR
help.s = help.s + "               Where 1 is the most active and 254 most inactive."+CR
help.s = help.s + "               255 disables APM, 128 is factory default"+CR
help.s = help.s + "  /SETAPM:n    Set the HDD APM level to this value and exit eeeHDD"+CR
help.s = help.s + "               This option is useful to find out the best APM level for your"+CR
help.s = help.s + "               drive. Best to be used from console!"+CR
help.s = help.s + "  /NOWARN      Do not display a warning on small APM values <100"+CR+CR


If Not IsUserAnAdmin_()
  MessageRequester("Error", "This program requires administrator rights."+Chr(13)+"Use 'surun' or 'runas' to start this program."+Chr(13)+"Quitting now.",#PB_MessageRequester_Ok)
  End  
EndIf


Procedure BlinkIcon()
  For i = 1 To 5
    ChangeSysTrayIcon(0, ImageID(1))
    Delay(300)
    ChangeSysTrayIcon(0, ImageID(0))
    Delay(300)
  Next i
EndProcedure

Procedure.s InfoText()
  txt.s = "APM value "+Str(APMValue)+" last set on "+FormatDate("%YYYY-%MM-%DD %HH:%II:%SS", Date())
  If tray=#True
    SetMenuItemText(0, 1, txt)
  EndIf
  ProcedureReturn txt
EndProcedure


Procedure setapm(APMValue.l)
  ; DO NOT USE THIS ANYMORE!
  ; NOTE: I leave this in because it works. But IOCTL_IDE_PASS_THROUGH does not work on Vista anymore!
  ; \\.\PhysicalDrive0

  ; hDevice = CreateFile_( "\\.\PhysicalDrive0", #GENERIC_READ | #GENERIC_WRITE, #FILE_SHARE_READ | #FILE_SHARE_WRITE, 0, #OPEN_EXISTING, 0, 0)
  ; If hDevice = #INVALID_HANDLE_VALUE
    ; PrintN( "CreateFile failed.")
    ; MessageRequester("Error", "Failed to open PhysicalDrive0. Ending now.",#PB_MessageRequester_Ok)
    ; End
  ; EndIf
  ; 
  ; ;Initialize buffers...
  ; ; INPUT Commandbuffer
  ; regBuf_IN.IDEREGS
  ; regBuf_IN\bFeaturesReg  = 0  ;0
  ; regBuf_IN\bSectorCountReg = 0  ;1 
  ; regBuf_IN\bSectorNumberReg  = 0  ;2
  ; regBuf_IN\bCylLowReg    = 0  ;3
  ; regBuf_IN\bCylHighReg   = 0  ;4
  ; regBuf_IN\bDriveHeadReg = 0  ;5
  ; regBuf_IN\bCommandReg   = 0  ;6
  ; regBuf_IN\bReserved   = 0  ; reg[7] is reserved for future use. Must be zero.
  ; regBuf_IN\DataBufferSize  = 0  ;8
  ; ; OUTPUT Commandbuffer
  ; regBuf_OUT.IDEREGS
  ; regBuf_OUT\bFeaturesReg = 0  ;0
  ; regBuf_OUT\bSectorCountReg  = 0  ;1 
  ; regBuf_OUT\bSectorNumberReg = 0  ;2
  ; regBuf_OUT\bCylLowReg   = 0  ;3
  ; regBuf_OUT\bCylHighReg  = 0  ;4
  ; regBuf_OUT\bDriveHeadReg  = 0  ;5
  ; regBuf_OUT\bCommandReg  = 0  ;6
  ; regBuf_OUT\bReserved    = 0  ; reg[7] is reserved for future use. Must be zero.
  ; regBuf_OUT\DataBufferSize = 0  ;8
  ; 
  ; bSize = SizeOf(regBuf_IN) ; Size of regBuf_IN - 8 for reg, 4 for DataBufferSize, 512 for Data
  ; 
  ; ; Prepare the ATA Command
  ; regBuf_IN\bCommandReg  = #WIN_SETFEATURES
  ; regBuf_IN\bFeaturesReg = #SETFEATURES_EN_APM
  ; If APMValue<1
    ; APMValue=1
  ; EndIf
  ; If APMValue>254
    ; APMValue = 255
    ; regBuf_IN\bFeaturesReg = #SETFEATURES_DIS_APM
    ; regBuf_IN\bSectorCountReg = 0
    ; ;PrintN("Disable APM.")
  ; Else
    ; regBuf_IN\bSectorCountReg = APMValue
    ; ;PrintN("Setting APM to "+ Str(APMValue)+".")
  ; EndIf
 ; 
  ; bytesRet.l = 0
  ; retval = DeviceIoControl_( hDevice, #IOCTL_IDE_PASS_THROUGH, @regBuf_IN, bSize, @regBuf_OUT, bSize, @bytesRet, #Null) 
  ; If retval=0
    ; CompilerIf #DEBUG = 1
      ; txt.s = "IOCTL_IDE_PASS_THROUGH failed. Error="+Str(GetLastError_())+Chr(13)
      ; txt.s = txt.s + "  In : CMD=0x"+ZHex(regBuf_IN\bCommandReg)+" FR=0x"+ZHex(regBuf_IN\bFeaturesReg)+" SC=0x"+ZHex(regBuf_IN\bSectorCountReg)+" SN=0x"+ZHex(regBuf_IN\bSectorNumberReg)+" CL=0x"+ZHex(regBuf_IN\bCylLowReg)+" CH=0x"+ZHex(regBuf_IN\bCylHighReg)+" SEL=0x"+ZHex(regBuf_IN\bDriveHeadReg)+Chr(13)
      ; txt.s = txt.s + " Out : CMD=0x"+ZHex(regBuf_OUT\bCommandReg)+" FR=0x"+ZHex(regBuf_OUT\bFeaturesReg)+" SC=0x"+ZHex(regBuf_OUT\bSectorCountReg)+" SN=0x"+ZHex(regBuf_OUT\bSectorNumberReg)+" CL=0x"+ZHex(regBuf_OUT\bCylLowReg)+" CH=0x"+ZHex(regBuf_OUT\bCylHighReg)+" SEL=0x"+ZHex(regBuf_OUT\bDriveHeadReg)+Chr(13)
    ; CompilerElse
      ; txt.s = ""
    ; CompilerEndIf
    ; MessageRequester("Error", "IOCTL_IDE_PASS_THROUGH failed. Reason:0x"+RSet(Hex(GetLastError_()), 8, "0")+Chr(13)+Chr(13)+txt,#PB_MessageRequester_Ok)
  ; EndIf
  ; ; PrintN( "bytesret: "+Str( bytesRet))
  ; ; PrintN("  In : CMD=0x"+ZHex(regBuf_IN\bCommandReg)+" FR=0x"+ZHex(regBuf_IN\bFeaturesReg)+" SC=0x"+ZHex(regBuf_IN\bSectorCountReg)+" SN=0x"+ZHex(regBuf_IN\bSectorNumberReg)+" CL=0x"+ZHex(regBuf_IN\bCylLowReg)+" CH=0x"+ZHex(regBuf_IN\bCylHighReg)+" SEL=0x"+ZHex(regBuf_IN\bDriveHeadReg))
  ; ; PrintN(" Out : CMD=0x"+ZHex(regBuf_OUT\bCommandReg)+" FR=0x"+ZHex(regBuf_OUT\bFeaturesReg)+" SC=0x"+ZHex(regBuf_OUT\bSectorCountReg)+" SN=0x"+ZHex(regBuf_OUT\bSectorNumberReg)+" CL=0x"+ZHex(regBuf_OUT\bCylLowReg)+" CH=0x"+ZHex(regBuf_OUT\bCylHighReg)+" SEL=0x"+ZHex(regBuf_OUT\bDriveHeadReg))
  ; 
  ; If hDevice
    ; CloseHandle_( hDevice )
  ; EndIf
EndProcedure

Procedure.l setataapm(APMValue.l)
  ; \\.\PhysicalDrive0
  hDevice = CreateFile_( "\\.\PhysicalDrive0", #GENERIC_READ | #GENERIC_WRITE, #FILE_SHARE_READ | #FILE_SHARE_WRITE, 0, #OPEN_EXISTING, 0, 0)
  If hDevice = #INVALID_HANDLE_VALUE
    PrintN( "CreateFile failed.")
    MessageRequester("Error", "Failed to open PhysicalDrive0. Ending now.",#PB_MessageRequester_Ok)
    End
  EndIf
  
  Define ab.ATA_PASS_THROUGH_EX_WITH_BUFFERS
  ab\length             = SizeOf(ATA_PASS_THROUGH_EX)
  ab\DataBufferOffset   = OffsetOf(ATA_PASS_THROUGH_EX_WITH_BUFFERS\ucDataBuf) 
  ab\TimeOutValue       = 1 ; 10
  dbsize.l              = SizeOf(ATA_PASS_THROUGH_EX_WITH_BUFFERS)

  ab\AtaFlags           = #ATA_FLAGS_DATA_IN
  ab\DataTransferLength = 512  
  ab\ucDataBuf[0]       = $CF   ; magic=0xcf
  ab\CurrentTaskFile[#inCommandReg]     = #WIN_SETFEATURES
  ab\CurrentTaskFile[#inFeaturesReg]    = #SETFEATURES_EN_APM
  
  If APMValue<1
    APMValue=1
  EndIf
  If APMValue>254
    APMValue=255
    ab\CurrentTaskFile[#inFeaturesReg]    = #SETFEATURES_DIS_APM
    ab\CurrentTaskFile[#inSectorCountReg] = 0
  Else
    ab\CurrentTaskFile[#inSectorCountReg] = APMValue
  EndIf
  
  If pdebug=#True
    PrintN(InfoText())
    PrintN("APMValue:            0x"+ZHex(APMValue)+" "+Str(APMValue))
    PrintN("Input registers:")
    PrintN("  inFeaturesReg      0x"+ZHex(ab\CurrentTaskFile[#inFeaturesReg]    )+"    inSectorCountReg  0x"+ZHex(ab\CurrentTaskFile[#inSectorCountReg]))
    PrintN("  inSectorNumberReg  0x"+ZHex(ab\CurrentTaskFile[#inSectorNumberReg])+"    inCylLowReg       0x"+ZHex(ab\CurrentTaskFile[#inCylLowReg]))
    PrintN("  inCylHighReg       0x"+ZHex(ab\CurrentTaskFile[#inCylHighReg]     )+"    inDriveHeadReg    0x"+ZHex(ab\CurrentTaskFile[#inDriveHeadReg]))
    PrintN("  inCommandReg       0x"+ZHex(ab\CurrentTaskFile[#inCommandReg])) ;+"    inSectorCountReg 0x"+ZHex(ab\CurrentTaskFile[#inSectorCountReg]))
  EndIf

  bytesRet.l = 0
  retval = DeviceIoControl_( hDevice, #IOCTL_ATA_PASS_THROUGH, @ab, dbsize.l, @ab, dbsize.l, @bytesRet, #Null) 
  If retval=0
    MessageRequester("Error", "IOCTL_ATA_PASS_THROUGH failed. Reason:0x"+RSet(Hex(GetLastError_()), 8, "0")+Chr(13)+Chr(13)+"Run eeeHDD with /DEBUG argument to get more information",#PB_MessageRequester_Ok)
  EndIf
  If pdebug=#True
    PrintN("Output registers:")
    PrintN("  outErrorReg        0x"+ZHex(ab\CurrentTaskFile[#outErrorReg]       )+"    outSectorCountReg 0x"+ZHex(ab\CurrentTaskFile[#outSectorCountReg]))
    PrintN("  outSectorNumberReg 0x"+ZHex(ab\CurrentTaskFile[#outSectorNumberReg])+"    outCylLowReg      0x"+ZHex(ab\CurrentTaskFile[#outCylLowReg]))
    PrintN("  outCylHighReg      0x"+ZHex(ab\CurrentTaskFile[#outCylHighReg]     )+"    outDriveHeadReg   0x"+ZHex(ab\CurrentTaskFile[#outDriveHeadReg]))
    PrintN("  outStatusReg       0x"+ZHex(ab\CurrentTaskFile[#outStatusReg])) ;+"    inSectorCountReg 0x"+ZHex(ab\CurrentTaskFile[#inSectorCountReg]))
    PrintN(CR)
    If retval=0
      PrintN("IOCTL_ATA_PASS_THROUGH failed. Reason:0x"+RSet(Hex(GetLastError_()), 8, "0"))
    EndIf
  EndIf
  
  If hDevice
    CloseHandle_( hDevice )
  EndIf
  If retval=0
    ProcedureReturn -1
  Else
    ProcedureReturn 0
  EndIf
EndProcedure

Procedure WinCallback(hwnd, msg, wParam, lParam)
  result = #PB_ProcessPureBasicEvents
  If msg = #WM_POWERBROADCAST
    Select wParam
    Case #PBT_APMQUERYSUSPEND
      If DisableSuspend = 1
        MessageBeep_(#MB_ICONASTERISK)
        ;ShowWindow_(hwnd, #SW_RESTORE)
        result = #BROADCAST_QUERY_DENY
      EndIf  
    Case #PBT_APMRESUMEAUTOMATIC
      ;MessageRequester("Information", "Resume from Suspend", #PB_MessageRequester_Ok)#
      setataapm(APMValue)
      InfoText()
      BlinkIcon()
    EndSelect
    
  EndIf
  ProcedureReturn result
EndProcedure

pcount = CountProgramParameters()
If pcount >0
  For i = 0 To pcount
    parm.s = UCase(ProgramParameter(i))
    If parm = "/DEBUG"
      pdebug=#True
    EndIf
    If parm = "/NOTRAY"
      tray=#False
    EndIf
    If parm = "/?"
      OpenConsole()
      Delay(100)
      PrintN(help.s)
      PrintN("-- PRESS RETURN --")
      Input()
      CloseConsole()
      End
    EndIf
    If parm = "/NOWARN"
      warnuser=#False
    EndIf
    If FindString(parm,"/SETAPM:",0)
      v.s = StringField(parm,2,":")
      av = Val(v)
      If av <1
        av=1
      EndIf
      If av>254
        av=255
      EndIf
      APMValue = av 
      OpenConsole()
      PrintN("eeeHDD Build:" + Str(#jaPBe_ExecuteBuild)+CR+CR)
      PrintN("SETAPM: Setting to "+Str(APMValue))
      If APMValue<100 And warnuser=#True
        PrintN(CR+CR+"!!WARNING!!   APMVALUE < 100   !!WARNING!!"+CR+CR)
        PrintN(help.s)
      EndIf
      
      If setataapm(APMValue)=0
        PrintN("Ok.")
      Else
        PrintN("Failed.")
      EndIf
      PrintN("-- PRESS RETURN --")
      Input()
      CloseConsole()
      End
    EndIf
    If FindString(parm,"/APMVALUE:",0)
      v.s = StringField(parm,2,":")
      av = Val(v)
      If av <1
        av=1
      EndIf
      If av>254
        av=255
      EndIf
      APMValue = av 
    EndIf
  Next
EndIf

If APMValue<100 And warnuser=#True
  MessageRequester("!!WARNING!!   APMVALUE < 100   !!WARNING!!", help.s , #MB_OK|#MB_ICONINFORMATION) 
EndIf

If pdebug = #True
  OpenConsole()
  PrintN("eeeHDD: Debug is turned on."+CR)
EndIf

If OpenWindow(0, 200, 300, 400, 200,"eeeHDD",#PB_Window_SystemMenu | #PB_Window_MinimizeGadget | #PB_Window_MaximizeGadget | #PB_Window_Invisible)
  SetWindowCallback(@WinCallback())

  If tray = #True
    If CreatePopupMenu(0)
      MenuItem(1, "")
      MenuBar()
      MenuItem(2, "Disable System Suspend")
      MenuItem(3, "Disable HDD APM now")
      MenuBar()
      MenuItem(4, "About / Help")
      MenuBar()
      MenuItem(5, "Quit")
      DisableMenuItem(0, 1, 1) ; Item 1 is used to display some text.
    EndIf 
    
    CatchImage(0, ?nicon)
    CatchImage(1, ?gicon)
    AddSysTrayIcon(0, WindowID(0), ImageID(0))
    SysTrayIconToolTip (0, "eeeHDD"+Chr(13)+"Build: " + Str(#jaPBe_ExecuteBuild) + Chr(13) + "APMValue: "+Str(APMValue))
  EndIf
  
  ; Set APM on startup
  setataapm(APMValue)

  If tray = #True
    InfoText()
    BlinkIcon()  
  EndIf
  
  Minimized=#True : quit.l=#False

  Repeat
    EventID.l=WaitWindowEvent()

    ; Make sure we check for System tray events...
    If EventID=#PB_Event_SysTray
      If EventType()=#PB_EventType_RightClick
        DisplayPopupMenu(0, WindowID(0))
      EndIf
    EndIf
  
    If EventID=#PB_Event_Menu ; <<-- PopUp Event
    ; Check for Right Mouse Button Menu and Restore Window or Quit...
      Select EventMenu()
        Case 2 ; Disable suspend
          If tray = 1
            If GetMenuItemState(0,2) = 1
              SetMenuItemState(0,2,0)
              DisableSuspend = 0
            Else
              SetMenuItemState(0,2,1)
              DisableSuspend = 1
            EndIf
          EndIf
        Case 3 ; Set HDD APM
          setataapm(APMValue)
          
          If tray = 1
            InfoText()
            BlinkIcon()
          EndIf
        Case 4 ; About/Help
          MessageRequester("About eeeHDD", help.s , #MB_OK|#MB_ICONINFORMATION)  
        Case 5 ; Quit
          quit=#True
      EndSelect
    EndIf
  
  
    If EventID=#PB_Event_CloseWindow And Minimized=#False ; <-- Check for the eXit Button and not SysTrayIcon
      ShowWindow_(WindowID(0),#SW_HIDE)
      Minimized=#True
    EndIf

  Until quit=#True
 
EndIf ;OpenWindow

End 



DataSection
  nicon: IncludeBinary "quiethd.ico"
  gicon: IncludeBinary "quiethd_gn.ico"
EndDataSection
   
; jaPBe Version=3.8.9.728
; Build=121
; ProductName=eeeHDD
; ProductVersion=1.0
; FileDescription=eeeHDD diables the Automatic Power Management (APM) of the primary Harddrive and eliminates the annoying click sound that the Seagate HDD's produces
; FileVersion=1.0 Build %build%
; InternalName=eeeHDD
; LegalCopyright=Freeware - This software comes with absolutely NO WARRANTY! Use it on your own risk!
; OriginalFilename=eeeHDD.exe
; EMail=joern.koerner@gmail.com
; Web=http://sites.google.com/site/eeehddsite/
; Language=0x0000 Language Neutral
; FirstLine=269
; CursorPosition=355
; EnableADMINISTRATOR
; EnableXP
; UseIcon=quiethd.ico
; ExecutableFormat=Windows
; Executable=C:\Dokumente und Einstellungen\injk\Eigene Dateien\Source\eeeHDD\trunk\eeeHDD.exe
; DontSaveDeclare
; EOF