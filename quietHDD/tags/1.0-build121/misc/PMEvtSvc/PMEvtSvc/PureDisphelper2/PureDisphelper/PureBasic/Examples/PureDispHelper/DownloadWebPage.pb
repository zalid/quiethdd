
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

If DownLoadWebPage("http://realsource.de", FileName)
  RunProgram(FileName)
EndIf

; IDE Options = PureBasic 4.20 Beta 2 (Windows - x86)
; CursorPosition = 34
; Folding = -
; EnableXP
; EnableUser
; HideErrorLog