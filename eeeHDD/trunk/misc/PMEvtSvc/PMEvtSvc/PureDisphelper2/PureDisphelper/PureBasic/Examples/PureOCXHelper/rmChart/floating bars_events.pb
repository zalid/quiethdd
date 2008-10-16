; example by ts-soft
; and event extensions by freak

EnableExplicit

XIncludeFile "..\ComEventSink.pb"

Procedure EventCallback(Event$, ParameterCount, *Params)
  Debug Event$

  If ParameterCount = 5
    Debug "  Button: " + Str(OCX_EventLong(*Params, 1))
    Debug "  X: " + Str(OCX_EventLong(*Params, 3))
    Debug "  Y: " + Str(OCX_EventLong(*Params, 4))
  EndIf

  Debug ""
EndProcedure

Define.l oRMChart
Define.s bars, tmp

Restore bars
Repeat
  Read tmp
  bars + tmp
Until tmp = ""

dhToggleExceptions(#True)

If OpenWindow(0, #PB_Ignore, #PB_Ignore, 550, 300, "RMChart loaded as string from Datasection") And CreateGadgetList(WindowID(0))

  oRMChart = OCX_CreateGadget(1, 0, 0, 550, 300, "RMChart.RMChartX")

  If oRMChart
    dhPutValue(oRMChart, ".RMCFile = %T", @bars)
    dhPutValue(oRMChart, ".RMCUserWatermark = %T", @"PureBasic")
    dhCallMethod(oRMChart, ".Draw")

    OCX_ConnectEvents(oRMChart, @EventCallback())
  EndIf

  While WaitWindowEvent() ! 16 : Wend

  CloseWindow(0)

  If oRMChart : dhReleaseObject(oRMChart) : EndIf

EndIf

DataSection
  bars:
  Data.s "00003550|00004300|000051|000073|00008-2894893|00009310|00011Tahoma|100011|100035|100045|10005-5|10006-5"
  Data.s "|1000911|100101|100111|100131|100181|100201|1002113|1002213|100238|100331|100341|100356|100378|100411"
  Data.s "|100468|100482|10052-16777216|10053-1120086|100544|100555|10056-16777216|10057-16777216|10060-16777216"
  Data.s "|10061-16777216|1006316|10064-5383962|100652|10066-16777011|10181Birth of a Killer App|10182Schedule*Reality"
  Data.s "|10187Design*Development*Testing*Bug Fixing*Documentation*Marketing|1020104/01*04/02*04/03*04/04*04/05*04"
  Data.s "/06*04/07*04/08*04/09*04/10*04/11*04/12*05/01|110011|110026|110044|110101|110131|11019-6751336|1102111|110221"
  Data.s "|1102312|110531*3*4*6*6*4*7*4*9*3*10*3|120011|120026|120044|120101|120132|12019-47872|1202111|120221|1202312"
  Data.s "|120531*.5*1.5*10.5*12*1*12*1*12.5*.5*2*11"
  Data.s ""
EndDataSection


; IDE Options = PureBasic v4.02 (Windows - x86)
; CursorPosition = 33
; FirstLine = 4
; Folding = -
; EnableXP
; HideErrorLog