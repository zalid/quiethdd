; example by kiffi
; enhanced by einander

#RMC_DATAAXISLEFT=1
#RMC_LINESTYLEDOT=2
#RMC_LABELAXISBOTTOM=8
#RMC_TEXTCENTER=0
#RMC_LINESTYLENONE=6
#RMC_BARSINGLE=1
#RMC_BAR_FLAT_GRADIENT2=3
#RMC_HATCHBRUSH_OFF=0
#RMC_BICOLOR_LABELAXIS=2
#RMC_CTRLSTYLEFLAT=0

Define.l oRMChart

dhToggleExceptions(#True)
;<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
hwnd= OpenWindow(0, #PB_Ignore, #PB_Ignore, 640, 450, "RMChart")
StickyWindow(0,1)
CreateGadgetList(hwnd)
oRMChart = OCX_CreateGadget(1, 0, 0, 640, 450, "RMChart.RMChartX")
Debug oRMChart
   
If oRMChart
    ; '************** Design the chart **********************
    dhCallMethod(oRMChart, ".Reset")
    dhPutValue(oRMChart, ".Font = %T", @"Comic Sans MS")
    dhPutValue(oRMChart, ".RMCBackColor = %u", -984833) ; 'AliceBlue'
    dhPutValue(oRMChart, ".RMCStyle = %d", #RMC_CTRLSTYLEFLAT)
    dhPutValue(oRMChart, ".RMCWidth = %d", 640)
    dhPutValue(oRMChart, ".RMCHeight = %d", 450)
   
    ; ************** Add Region 1 *****************************
   
    dhCallMethod(oRMChart, ".AddRegion")
   
    dhPutValue(oRMChart, ".Region(%d).Left = %u", 1, 5)
    dhPutValue(oRMChart, ".Region(%d).Top = %u", 1, 5)
    dhPutValue(oRMChart, ".Region(%d).Width = %u", 1, -5)
    dhPutValue(oRMChart, ".Region(%d).Height = %u", 1, -5)
    dhPutValue(oRMChart, ".Region(%d).Footer = %T", 1, @"")
   
    ; ************** Add caption to region 1 *******************
   
    dhCallMethod(oRMChart, ".Region(%d).AddCaption", 1)
   
    dhPutValue(oRMChart, ".Region(%d).Caption.Titel = %T", 1, @"Press right button and move mouse")
    dhPutValue(oRMChart, ".Region(%d).Caption.BackColor = %u", 1, -16776961) ; 'Blue'
    dhPutValue(oRMChart, ".Region(%d).Caption.TextColor = %u", 1, -256) ; 'Yellow'
    dhPutValue(oRMChart, ".Region(%d).Caption.FontSize = %d", 1, 11)
    dhPutValue(oRMChart, ".Region(%d).Caption.Bold = %b", 1, #True)
   
    ; ************** Add grid to region 1 *****************************
   
    dhCallMethod(oRMChart, ".Region(%d).AddGrid", 1)
   
    dhPutValue(oRMChart, ".Region(%d).Grid.BackColor = %u", 1, -657956) ; 'Beige'
    dhPutValue(oRMChart, ".Region(%d).Grid.AsGradient = %b", 1, #False)
    dhPutValue(oRMChart, ".Region(%d).Grid.BicolorMode = %d", 1, #RMC_BICOLOR_LABELAXIS)
    dhPutValue(oRMChart, ".Region(%d).Grid.Left = %d", 1, 0)
    dhPutValue(oRMChart, ".Region(%d).Grid.Top = %d", 1, 0)
    dhPutValue(oRMChart, ".Region(%d).Grid.Width = %d", 1, 0)
    dhPutValue(oRMChart, ".Region(%d).Grid.Height = %d", 1, 0)
   
    ; ************** Add data axis to region 1 *****************************
   
    dhCallMethod(oRMChart, ".Region(%d).AddDataAxis", 1)
   
    dhPutValue(oRMChart, ".Region(%d).DataAxis(%d).Alignment = %d", 1, 1, #RMC_DATAAXISLEFT)
    dhPutValue(oRMChart, ".Region(%d).DataAxis(%d).MinValue = %d", 1, 1, 0)
    dhPutValue(oRMChart, ".Region(%d).DataAxis(%d).MaxValue = %d", 1, 1, 100)
    dhPutValue(oRMChart, ".Region(%d).DataAxis(%d).TickCount = %d", 1, 1, 11)
    dhPutValue(oRMChart, ".Region(%d).DataAxis(%d).FontSize = %d", 1, 1, 8)
    dhPutValue(oRMChart, ".Region(%d).DataAxis(%d).TextColor = %u", 1, 1, -16777216) ;  'Black'
    dhPutValue(oRMChart, ".Region(%d).DataAxis(%d).LineColor = %u", 1, 1, -16777216) ;  'Black'
    dhPutValue(oRMChart, ".Region(%d).DataAxis(%d).LineStyle = %d", 1, 1, #RMC_LINESTYLEDOT)
    dhPutValue(oRMChart, ".Region(%d).DataAxis(%d).DecimalDigits = %d", 1, 1, 0)
    dhPutValue(oRMChart, ".Region(%d).DataAxis(%d).AxisUnit = %T", 1, 1, @"")
    dhPutValue(oRMChart, ".Region(%d).DataAxis(%d).AxisText = %T", 1, 1, @"")
   
    ; ************** Add label axis to region 1 *****************************
   
    dhCallMethod(oRMChart, ".Region(%d).AddLabelAxis", 1)
   
    dhPutValue(oRMChart, ".Region(%d).LabelAxis.AxisCount = %d", 1, 1)
    dhPutValue(oRMChart, ".Region(%d).LabelAxis.TickCount = %d", 1, 5)
    dhPutValue(oRMChart, ".Region(%d).LabelAxis.Alignment = %d", 1, #RMC_LABELAXISBOTTOM)
    dhPutValue(oRMChart, ".Region(%d).LabelAxis.FontSize = %d", 1, 8)
    dhPutValue(oRMChart, ".Region(%d).LabelAxis.TextColor = %u", 1, -16777216) ;  'Black'
    dhPutValue(oRMChart, ".Region(%d).LabelAxis.TextAlignment = %d", 1, #RMC_TEXTCENTER)
    dhPutValue(oRMChart, ".Region(%d).LabelAxis.LineColor = %u", 1, -16777216) ;  'Black'
    dhPutValue(oRMChart, ".Region(%d).LabelAxis.LineStyle = %d", 1, #RMC_LINESTYLENONE)
    dhPutValue(oRMChart, ".Region(%d).LabelAxis.LabelString = %T", 1, @"Label 1*Label 2*Label 3*Label 4*Label 5")
   
    ; ;************** Add Series 1 to region 1 *******************************
    ;
    dhCallMethod(oRMChart, ".Region(%d).AddBarSeries", 1)
   
    dhPutValue(oRMChart, ".Region(%d).BarSeries(%d).SeriesType = %d", 1, 1, #RMC_BARSINGLE)
    dhPutValue(oRMChart, ".Region(%d).BarSeries(%d).SeriesStyle = %d", 1, 1, #RMC_BAR_FLAT_GRADIENT2)
    dhPutValue(oRMChart, ".Region(%d).BarSeries(%d).Lucent = %b", 1, 1, #False)
    dhPutValue(oRMChart, ".Region(%d).BarSeries(%d).Color = %u", 1, 1, -10185235) ; 'CornflowerBlue'
    dhPutValue(oRMChart, ".Region(%d).BarSeries(%d).Horizontal = %b", 1, 1, #False)
    dhPutValue(oRMChart, ".Region(%d).BarSeries(%d).WhichDataAxis = %d", 1, 1, 1)
    dhPutValue(oRMChart, ".Region(%d).BarSeries(%d).ValueLabelOn = %b", 1, 1, #False)
    dhPutValue(oRMChart, ".Region(%d).BarSeries(%d).PointsPerColumn = %d", 1, 1, 1)
    dhPutValue(oRMChart, ".Region(%d).BarSeries(%d).HatchMode = %d", 1, 1, #RMC_HATCHBRUSH_OFF)
    dhPutValue(oRMChart, ".Region(%d).BarSeries(%d).DataString = %T", 1, 1, @"50*70*40*60*30")
    dhCallMethod(oRMChart, ".Draw")
EndIf

#SEP="*"
Repeat
    EV=WaitWindowEvent() 
    _DX=DesktopMouseX()
    _DY=DesktopMouseY()
   
    _MK=Abs(GetAsyncKeyState_(#VK_LBUTTON) +GetAsyncKeyState_(#VK_RBUTTON)*2+GetAsyncKeyState_(#VK_MBUTTON)*3)/$8000   
    If _MK=2
        A$=Str(_DX / 10)
        B$=Str(_DY / 10)
        c$=Str((_DX+_DY)/10)
        d$=Str((_DX+_DY)/7)
        e$=Str((_DX+_DY)/20)
        Value$=A$+#SEP+B$+#SEP+c$+#SEP+d$+#SEP+e$+#SEP
        dhPutValue(oRMChart, ".Region(%d).BarSeries(%d).DataString = %T", 1, 1, @Value$)
        dhCallMethod(oRMChart, ".Draw")
    EndIf
Until EV=#WM_CLOSE 
   
If oRMChart : dhReleaseObject(oRMChart) : EndIf
CloseWindow(0)
End 
; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 126
; FirstLine = 80
; Folding = -
; EnableXP
; HideErrorLog