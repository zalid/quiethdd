; c example from disphelper
; adapted to PB by ts-soft

EnableExplicit

dhToggleExceptions(#True)

Procedure.s Replace(szPattern.s, szString.s, szReplacement.s, bIgnoreCase.l = #True)
  Protected regEx.l, Result.l, szResult.s

  If bIgnoreCase <> 0 : bIgnoreCase = 1 : EndIf

  regEx = dhCreateObject("VBScript.RegExp")

  If regEx
    dhPutValue(regEx, ".Pattern = %T", @szPattern)
    dhPutValue(regEx, ".IgnoreCase = %b", bIgnoreCase)
    dhPutValue(regEx, ".Global = %b", #True)

    dhGetValue("%T", @Result, regEx, ".Replace (%T,%T)", @szString, @szReplacement)

    If Result
      szResult = PeekS(Result)
      dhFreeString(Result)
    EndIf

    dhReleaseObject(regEx)
  EndIf
  ProcedureReturn szResult
EndProcedure

Debug Replace("fox", "The quick brown fox jumped over the lazy dog.", "cat")

; IDE Options = PureBasic 4.20 Beta 2 (Windows - x86)
; Folding = -
; EnableXP
; EnableUser
; HideErrorLog