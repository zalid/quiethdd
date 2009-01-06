; based on excel example by kiffi and mk-soft
; example for openoffice by KIKI
; small corrections by ts-soft

EnableExplicit

Define DISPHELPER_NO_OCX_CreateGadget

XIncludeFile "DispHelper_Include.pb"
XIncludeFile "VariantHelper_Include.pb"

Define.l oSM, oDesk, oDoc, oTable, oCell

Procedure ChangeCellText(Table, Cell.s, Value.s)
  Protected oCell.l
  dhGetValue("%o", @oCell, Table, ".GetCellRangeByName(%T", @Cell)
  dhPutValue(oCell, ".String = %T", @Value)
  dhReleaseObject(oCell)
EndProcedure

Procedure PutCellValue(Table, Cell.s, *Value.VARIANT)
  Protected oCell.l
  dhGetValue("%o", @oCell, Table, ".GetCellRangeByName(%T", @Cell)
  dhPutValue(oCell, ".value = %v", *Value)
  dhReleaseObject(oCell)
EndProcedure

dhInitializeImp()

Define.safearray *openpar
Define.variant openarray
 

V_ARRAY_DISP(openarray) = *openpar

dhToggleExceptions(#True)

Define.variant wert1, wert2, result, text
Define.d wert3 = 5.55555555555555

V_DOUBLE(wert1) = 3.33333333333333
V_DOUBLE(wert2) = 4.44444444444444
V_STR(text) = T_BSTR("Hello World")


oSM = dhCreateObject("com.sun.star.ServiceManager")
If oSM
  dhGetValue("%o",@oDesk,oSM, ".CreateInstance(%T)", @"com.sun.star.frame.Desktop")
  dhGetValue("%o",@oDoc, oDesk,".LoadComponentFromURL(%T,%T,%d,%v)",@"private:factory/scalc", @"_blank", 0, openarray)
  dhGetValue("%o",@oTable, odoc ,".Sheets.GetByIndex(%d)",  0)

  ChangeCellText(oTable, "A1", "Feel")
  ChangeCellText(oTable, "B1", "the")
  ChangeCellText(oTable, "C1", "pure")
  ChangeCellText(oTable, "D1", "Power")
  ChangeCellText(oTable, "A2", "The")
  ChangeCellText(oTable, "A3", "Pure")
  ChangeCellText(oTable, "A4", "Power")
  
  ; with VARIANT
  PutCellValue(oTable, "B2", wert1)
  PutCellValue(oTable, "B3", wert2)

  ; with PB-Double
  dhGetValue("%o", @oCell, oTable, ".GetCellRangeByName(%T", @"B4")
  dhPutValue(oCell, ".value = %e", @wert3)
  dhReleaseObject(oCell)
  
  dhGetValue("%o", @oCell, oTable, ".GetCellRangeByName(%T", @"C2")
  dhPutValue(oCell, ".String = %v", text)
  dhReleaseObject(oCell)

  dhGetValue("%o", @oCell, oTable, ".GetCellRangeByName(%T", @"B2")
  dhGetValue("%v", @result, oCell, ".Value")
  dhReleaseObject(oCell)

  MessageRequester("PureDispHelper-OpenCalcDemo", "Result Cells(2,2): " + VT_STR(result))
  MessageRequester("PureDispHelper-OpenCalcDemo", "Click OK to close Open Office")

  dhCallMethod(odoc,".close(%b)",#True)
  dhReleaseObject(oTable)
  dhReleaseObject(oDoc)
  dhReleaseObject(oSM)
EndIf
dhUninitialize()
End

; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 28
; FirstLine = 3
; Folding = -
; EnableXP
; HideErrorLog