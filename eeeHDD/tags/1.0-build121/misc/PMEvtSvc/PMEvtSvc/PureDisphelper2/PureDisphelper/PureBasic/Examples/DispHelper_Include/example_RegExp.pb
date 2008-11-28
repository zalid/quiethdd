; /* --
; regexp.c:
;   Demonstrates using the VBScript.RegExp object To provide support For regular
; expressions.
;  -- */

; changed from c to pb by ts-soft

EnableExplicit

Define DISPHELPER_NO_OCX_CreateGadget ; we don't use it
XIncludeFile "Disphelper_Include.pb"
;
; /* **************************************************************************
;  * RegExp:
;  *   Executes a regular expression And prints out matches.
;  *
;  ============================================================================ */
Procedure RegExp(szPattern.s, szString.s, bIgnoreCase.l = #True)
  Protected regEx.l = dhCreateObject("VBScript.RegExp")
  Protected match.l, nFirstIndex.l, nLength.l, szValue.l
  
  If regEx
    dhPutValue(regEx, ".Pattern = %T", @szPattern)
    dhPutValue(regEx, ".IgnoreCase = %b", bIgnoreCase)
    dhPutValue(regEx, ".Global = %b", #True)
    
    FOR_EACH(match, regEx, ".Execute(%T)", @szString)
    
      nFirstIndex = 0 : nLength = 0 : szValue = 0
      
      dhGetValue("%d", @nFirstIndex, match, ".FirstIndex")
      dhGetValue("%d", @nLength, match, ".Length")
      dhGetValue("%T", @szValue, match, ".Value")
      
      If szValue
        Debug "Match found at charachters " + Str(nFirstIndex) + " to " + Str(nFirstIndex + nLength) + " Match value is " + PeekS(szValue)
        dhFreeString(szValue)
      EndIf
      
    NEXT_(match)
    
    dhReleaseObject(regEx)
  EndIf
EndProcedure
; 
; 
; /* **************************************************************************
;  * Replace:
;  *   Performs a regular expression replace And prints out the result.
;  *
;  ============================================================================ */
Procedure Replace(szPattern.s, szString.s, szReplacement.s, bIgnoreCase.l = #True)
  Protected regEx.l = dhCreateObject("VBScript.RegExp")
  Protected szResult.l
  
  If regEx
    dhPutValue(regEx, ".Pattern = %T", @szPattern)
    dhPutValue(regEx, ".IgnoreCase = %b", bIgnoreCase)
    dhPutValue(regEx, ".Global = %b", #True)
    
    dhGetValue("%T", @szResult, regEx, ".Replace(%T, %T)", @szString, @szReplacement)
    
    If szResult <> 0
      Debug PeekS(szResult)
      dhFreeString(szResult)
    EndIf
  EndIf
EndProcedure

dhInitializeImp()
dhToggleExceptions(#True)

Debug "Running regular expression find sample..."
Debug "------------------------------------------"
RegExp("is.", "IS1 is2 IS3 is4", #True)

Debug ""
Debug "Running regular expression replace sample..."
Debug "------------------------------------------"
Replace("fox", "The quick brown fox jumped over the lazy dog.", "cat", #True)

dhUninitialize()
; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 39
; FirstLine = 18
; Folding = -
; EnableXP
; HideErrorLog