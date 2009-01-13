EnableExplicit

IncludeFile "VariantHelper_Include.pb"

;- Test SafeArray

Define.safearray *psa, *psa_dest

Define.variant var, var_dest

Define.l i


; Create SafeArray
Debug "Create SafeArray with first element is number 1 and 10 elements as string"
*psa = saCreateSafeArray(#TString, 1, 10)

Debug "First Element: " + Str(saLBound(*psa))
Debug "Last Element: " + Str(saUBound(*psa))
Debug "Count Element: " + Str(saCount(*psa))

; Fill SafeArray
For i = saLBound(*psa) To saUBound(*psa)
    SA_BSTR(*psa, i) = T_BSTR("Number " + Str(i + 100))
Next

; Output SafeArray
If *psa
  For i = saLBound(*psa) To saUBound(*psa)
    Debug "Index: " + Str(i) + " Value: " + SA_STR(*psa, i)
  Next
EndIf

; Free SafeArray
If Not saFreeSafeArray(*psa)
  Debug saGetLastMessage()
EndIf


; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 15
; Folding = -