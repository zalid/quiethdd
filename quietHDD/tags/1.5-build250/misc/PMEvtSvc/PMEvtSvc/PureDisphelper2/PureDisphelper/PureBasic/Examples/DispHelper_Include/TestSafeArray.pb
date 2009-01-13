EnableExplicit

IncludeFile "VariantHelper_Include.pb"

;- Test SafeArray

Define.safearray *psa, *psa_dest

Define.variant var, var_dest

Define.l i

; Create SafeArray
Debug "Create SafeArray with first element is number 10 and 100 elements as long"
*psa = saCreateSafeArray(#TLong, 10, 100)

; Einzelnen Wert zuweisen
SA_LONG(*psa, 3) = 9999

; Liste füllen

Debug "First Element: " + Str(saLBound(*psa))
Debug "Last Element: " + Str(saUBound(*psa))
Debug "Count Element: " + Str(saCount(*psa))

; Fill SafeArray
For i = saLBound(*psa) To saUBound(*psa)
    SA_LONG(*psa, i) = i + 100
Next

; Set Single Value
SA_LONG(*psa, saLBound(*psa) + 13) = 9999


; Set Variant SafeArray as Long
V_ARRAY_LONG(var) = *psa

; Copy Variant for testing
VariantCopy_(var_dest, var)

; Get SafeArray from Variant
*psa_dest = VT_ARRAY(var_dest)

; Output SafeArray
If *psa_dest
  For i = saLBound(*psa_dest) To saUBound(*psa_dest)
    Debug "Index: " + Str(i) + " Value: " + Str(SA_LONG(*psa_dest, i))
  Next
EndIf

; Free Variant - destroy BSTR and SafeArray destroy automatic
V_EMPTY(var)
V_EMPTY(var_dest)


; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 16
; Folding = -