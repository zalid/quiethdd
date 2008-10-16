
EnableExplicit

dhToggleExceptions(#True)

Define.l oWord = dhCreateObject("Word.Application")

If oWord

  dhPutValue      (oWord, "Visible = %b", #True)
  dhCallMethod    (oWord, "Documents.Add")
  dhCallMethod    (oWord, "Selection.TypeText(%s)", @"PureDispHelper Sample")
  dhReleaseObject (oWord) : oWord = 0

EndIf
