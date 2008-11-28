
EnableExplicit

dhToggleExceptions(#True)

Define.l oReport, I, sum.l, tmp.s

If OpenWindow(0, #PB_Ignore, #PB_Ignore, 640, 480, "Report-Generator")
  If CreateGadgetList(WindowID(0))
    oReport = OCX_CreateGadget(#PB_Any, 0, 0, 0, 0, "CatchysoftReport.Report")
    If oReport
      dhCallMethod(oReport, "InitReport")
      dhPutValue(oReport, "ReportName = %T", @"Simple Report")
      dhPutValue(oReport, "ColumnName = %T", @"Column 1")
      dhPutValue(oReport, "ColumnName = %T", @"Column 2")
      dhPutValue(oReport, "ColumnName = %T", @"Column 3")
      dhPutValue(oReport, "ColumnWidth = %d", 20)
      dhPutValue(oReport, "ColumnWidth = %d", 20)
      dhPutValue(oReport, "ColumnWidth = %d", 60)
      
      For I = 1 To 100
        tmp.s = Str(I)
        dhPutValue(oReport, "FieldText = %T", @tmp)
        dhPutValue(oReport, "FieldText = %T", @"test")
        dhPutValue(oReport, "FieldText = %T", @"10")
        sum + 10
      Next
      
      dhPutValue(oReport, "Summary = %T", @"")
      dhPutValue(oReport, "Summary = %T", @"Sum")
      tmp = Str(sum)
      dhPutValue(oReport, "Summary = %T", @tmp)
      
      dhCallMethod(oReport, "PrintPreview")
    Else
      MessageRequester("CatchysoftReport", "Pleas download the required Software at:" + #LF$ + "http://www.catchysoft.com/reportgeneratord.html")
    EndIf
  EndIf
  
  While WaitWindowEvent() ! 16 : Wend
  If oReport : dhReleaseObject(oReport) : EndIf
  CloseWindow(0)
EndIf

; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 20
; Folding = -
; EnableXP
; DisableDebugger
; HideErrorLog