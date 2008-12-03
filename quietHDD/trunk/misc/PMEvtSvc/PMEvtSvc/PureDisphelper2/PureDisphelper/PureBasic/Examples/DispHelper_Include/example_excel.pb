; example by Kiffi
; enhanced by mk-soft

EnableExplicit

Define DISPHELPER_NO_FOR_EACH
Define DISPHELPER_NO_OCX_CreateGadget

XIncludeFile "DispHelper_Include.pb"
XIncludeFile "VariantHelper_Include.pb"

Define.l ExcelApp, Workbook

dhInitializeImp()
dhToggleExceptions(#True)

ExcelApp = dhCreateObject("Excel.Application")

Define.variant result
Define.d value1, value2, value3, date, retval
Define.w short_value

value1 = 100.5
value2 = 10.1
value3 = 99.9
date = T_DATE(Date())
short_value = 32767

If ExcelApp
 
  dhPutValue(ExcelApp, ".Visible = %b", #True)
 
  dhGetValue("%o", @Workbook, ExcelApp, ".Workbooks.Add")
 
  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %T", 1, 1, @"Feel")
  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %T", 1, 2, @"the")
  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %T", 1, 3, @"pure")
  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %T", 1, 4, @"Power")
 
  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %T", 2, 1, @"Value1")
  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %T", 3, 1, @"Value2")
  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %T", 4, 1, @"Value3")
  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %T", 5, 1, @"Now")
  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %T", 6, 1, @"Short")
 
 
  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %e", 2, 2, @value1)
  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %e", 3, 2, @value2)
  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %e", 4, 2, @value3)
  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %D", 5, 2, @date)
  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %i", 6, 2, short_value)

  dhGetValue("%e", @retval, ExcelApp, "Cells(%d, %d).Value", 4, 2)
  dhGetValue("%v", @result, ExcelApp, "Cells(%d, %d).Value", 5, 2)
 
  MessageRequester("PureDispHelper-ExcelDemo", "Result Cells(4,2) (Value 3): " + StrD(retval))
  MessageRequester("PureDispHelper-ExcelDemo", "Result Cells(5,2) (Date): " + VT_STR(result))
 
  MessageRequester("PureDispHelper-ExcelDemo", "Click OK to close Excel")
 
  dhCallMethod(ExcelApp, ".Quit")
 
  dhReleaseObject(Workbook) : Workbook = 0
  dhReleaseObject(ExcelApp) : ExcelApp = 0
 
Else
 
  MessageRequester("PureDispHelper-ExcelDemo", "Couldn't create Excel-Object")
 
EndIf

dhUninitialize()

; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 72
; Folding = -
; EnableXP
; HideErrorLog