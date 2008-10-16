IncludeFile "wmi.pb"

 OpenConsole()
 Print("Press RETURN to quit.")

 WMI_INIT()
 WMI_Call("Select * FROM Win32_PowerManagementEvent", "") ;, "Drive, Caption, Description, DeviceID")
 ResetList(wmidata())
 
 While Inkey()=""

 While NextElement(wmidata())
   Debug wmidata()  ; Alle Listenelemente darstellen / show all elements
 Wend
 
 Wend
WMI_RELEASE("OK") 


; IDE Options = PureBasic 4.20 (Windows - x86)
; CursorPosition = 2
; Folding = -
; UseMainFile = PMEvent.pb
; Executable = PMEvent.exe