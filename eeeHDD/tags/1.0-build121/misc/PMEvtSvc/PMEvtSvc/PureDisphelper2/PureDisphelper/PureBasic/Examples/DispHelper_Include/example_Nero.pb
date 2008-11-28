; Nero example
; Debug all CD/DVD Devices
; by DataMiner and ts-soft

EnableExplicit

Define DISPHELPER_NO_OCX_CreateGadget

XIncludeFile "DispHelper_Include.pb"

dhInitializeImp()

Structure DriveInfo
  DeviceName.s
  HostAdapterName.s
  DriveLetter.s
EndStructure

dhToggleExceptions(#True)

Define obj.l, objNeroDrives.l, objNeroDriveItem.l, szResult.l, Result.l, i.l

NewList DI.DriveInfo()

obj = dhCreateObject("Nero.Nero")

If obj

  dhGetValue("%o", @objNeroDrives, obj, "GetDrives(%d)", 0)
 
  If objNeroDrives
    dhGetValue("%d", @Result, objNeroDrives, "Count()")

    If Result 
      For i = 0 To Result - 1
        dhGetValue("%o", @objNeroDriveItem, objNeroDrives, "Item(%d)", i)
        If objNeroDriveItem
          AddElement(DI())
          dhGetValue("%T", @szResult, objNeroDriveItem, "DeviceName")
          If szResult
            DI()\DeviceName = PeekS(szResult)
            dhFreeString(szResult)
          EndIf
          dhGetValue("%T", @szResult, objNeroDriveItem, "HostAdapterName")
          If szResult
            DI()\HostAdapterName = PeekS(szResult)
            dhFreeString(szResult)
          EndIf
          dhGetValue("%T", @szResult, objNeroDriveItem, "DriveLetter")
          If szResult
            DI()\DriveLetter = PeekS(szResult)
            dhFreeString(szResult)
          EndIf
          dhReleaseObject(objNeroDriveItem)
        EndIf
      Next
    EndIf
    dhReleaseObject(objNeroDrives)
  EndIf
  dhReleaseObject(obj)
EndIf

dhUninitialize()

ForEach DI()
  With DI()
    Debug \DriveLetter + ":    " + \HostAdapterName + "        " + \DeviceName
  EndWith
Next
; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 53
; FirstLine = 18
; Folding = -
; EnableXP
; HideErrorLog