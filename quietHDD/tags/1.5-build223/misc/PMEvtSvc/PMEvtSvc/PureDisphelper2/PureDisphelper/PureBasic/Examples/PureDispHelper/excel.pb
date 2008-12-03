; example by Kiffi

EnableExplicit

Define.l ExcelApp, Workbook

dhToggleExceptions(#True)

ExcelApp = dhCreateObject("Excel.Application")

If ExcelApp

  dhPutValue(ExcelApp, ".Visible = %b", #True)

  dhGetValue("%o", @Workbook, ExcelApp, ".Workbooks.Add")

  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %s", 1, 1, @"Feel")
  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %s", 2, 1, @"the")
  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %s", 3, 1, @"pure")
  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %s", 4, 1, @"Power")

  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %s", 1, 2, @"the")
  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %s", 1, 3, @"pure")
  dhPutValue(ExcelApp, "Cells(%d, %d).Value = %s", 1, 4, @"Power")

  MessageRequester("PureDispHelper-ExcelDemo", "Click OK to close Excel")

  dhCallMethod(ExcelApp, ".Quit")

  dhReleaseObject(Workbook) : Workbook = 0
  dhReleaseObject(ExcelApp) : ExcelApp = 0

Else

  MessageRequester("PureDispHelper-ExcelDemo", "Couldn't create Excel-Object")

EndIf
