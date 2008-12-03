
Define DISPHELPER_NO_FOR_EACH
Define DISPHELPER_NO_OCX_CreateGadget

XIncludeFile "DispHelper_Include.pb"; if you use the userlib, comment this out

dhInitializeImp(); if you use the userlib, comment this out

dhToggleExceptions(#True)

Procedure DownLoadWebPage(szURL.s, szFileName.s)
  Protected objHTTP.l, Response.l, Status.l, File.l

  objHTTP = dhCreateObject("MSXML2.XMLHTTP")
  If objHTTP
    dhCallMethod(objHTTP, ".Open(%T, %T, %b)", @"GET", @szURL, #False)
    dhCallMethod(objHTTP, ".Send")

    dhGetValue("%T", @Status, objHTTP, ".StatusText")
    If Status <> 0
      If UCase(PeekS(Status)) = "OK"
        dhFreeString(Status)
        dhGetValue("%T", @Response, objHTTP, ".ResponseText")
        If Response <> 0
          File = CreateFile(#PB_Any, szFileName)
          If File
            WriteString(File, PeekS(Response))
            CloseFile(File)
            dhFreeString(Response)
            ProcedureReturn #True
          EndIf
          dhFreeString(Response)
        EndIf
      EndIf
    EndIf
    dhReleaseObject(objHTTP)
  EndIf
EndProcedure

Define.s FileName = GetTemporaryDirectory() + "test.html"

If DownLoadWebPage("http://ts-soft.eu", FileName)
  RunProgram(FileName)
EndIf

dhUninitialize(); if you use the userlib, comment this out

; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 1
; Folding = -
; EnableUnicode
; EnableXP
; HideErrorLog