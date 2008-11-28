EnableExplicit

dhToggleExceptions(#True)

Procedure FileExists(file.s)
  Protected oFSO.l, Result.l

  oFSO = dhCreateObject("Scripting.FileSystemObject")
  If oFSO
    dhGetValue("b%", @Result, oFSO, "FileExists(%T)", @file)
    dhReleaseObject(oFSO)
    If Result = 0
      ProcedureReturn #False
    Else
      ProcedureReturn #True
    EndIf
  EndIf
EndProcedure

Procedure FolderExists(folder.s)
  Protected oFSO.l, Result.l

  oFSO = dhCreateObject("Scripting.FileSystemObject")
  If oFSO
    dhGetValue("b%", @Result, oFSO, "FolderExists(%T)", @folder)
    dhReleaseObject(oFSO)
    If Result = 0
      ProcedureReturn #False
    Else
      ProcedureReturn #True
    EndIf
  EndIf
EndProcedure

Procedure.s GetTempName()
  Protected oFSO.l, Result.l, szResult.s

  oFSO = dhCreateObject("Scripting.FileSystemObject")
  If oFSO
    dhGetValue("T%", @Result, oFSO, "GetTempName")
    If Result
      szResult = PeekS(Result)
      dhFreeString(Result)
    EndIf
    dhReleaseObject(oFSO)
  EndIf
  ProcedureReturn szResult
EndProcedure

Procedure.s GetBaseName(file.s)
  Protected oFSO.l, Result.l, szResult.s

  oFSO = dhCreateObject("Scripting.FileSystemObject")
  If oFSO
    dhGetValue("T%", @Result, oFSO, "GetBaseName(%T)", @file)

    If Result
      szResult = PeekS(Result)
      dhFreeString(Result)
    EndIf
    dhReleaseObject(oFSO)
  EndIf
  ProcedureReturn szResult
EndProcedure

Debug FileExists("c:\autoexec.bat")
Debug FolderExists("c:\windows")
Debug GetTempName()
Debug GetBaseName(ProgramFilename())
