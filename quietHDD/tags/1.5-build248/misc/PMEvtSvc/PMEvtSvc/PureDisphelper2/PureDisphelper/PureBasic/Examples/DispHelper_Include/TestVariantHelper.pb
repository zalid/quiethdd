
IncludeFile "VariantHelper_Include.pb"

;-Test

Define.variant value, value2

; Wert ByRef
wert = 1000
V_LONG_BYREF(value) = @wert
Debug VT_STR(value)

; Double ByRef
dVal.d = 1000.0
V_DOUBLE_BYREF(value) = @dVal
Debug VT_STR(value)

datum.d = T_DATE(Date())

; Test Date
V_DATE(value) = datum
Debug VT_STR(value)
Debug FormatDate("%dd.%mm.%yyyy %hh:%ii:%ss", VT_DATE(value))

; Test Date ByRef
V_DATE_BYREF(value) = @datum
Debug VT_STR(value)
Debug FormatDate("%dd.%mm.%yyyy %hh:%ii:%ss", VT_DATE(value))

; Test LONG to Date
V_LONG(value) = datum
Debug VT_STR(value)
Debug FormatDate("%dd.%mm.%yyyy %hh:%ii:%ss", VT_DATE(value))

; Test Boolean True
V_BOOL(value) = T_BOOL(#True)
Debug VT_STR(value)
Debug VT_BOOL(value)

; Test Boolean False
V_BOOL(value) = T_BOOL(#False)
Debug VT_STR(value)
Debug VT_BOOL(value)

; Test Boolean True ByRef

result = T_BOOL(#True)
V_BOOL_BYREF(value) = @result
Debug VT_STR(value)
Debug VT_BOOL(value)

; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 1
; Folding = -