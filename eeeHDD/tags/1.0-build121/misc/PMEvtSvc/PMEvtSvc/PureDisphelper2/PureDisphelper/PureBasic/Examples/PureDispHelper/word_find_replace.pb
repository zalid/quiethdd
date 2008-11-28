; example by NUM3

dhToggleExceptions(#True)

Define.l oWord = dhCreateObject("Word.Application")

Define.s text = "PureDispHelper Sample" + #CRLF$
Define.l wdReplaceAll = 2

If oWord

  dhPutValue      (oWord, "Visible = %b", #True)

  dhCallMethod    (oWord, "Documents.Add")

  dhCallMethod    (oWord, "Selection.TypeText(%s)", @text)
  dhCallMethod    (oWord, "Selection.TypeText(%s)", @text)
  dhCallMethod    (oWord, "Selection.TypeText(%s)", @text)
  dhCallMethod    (oWord, "Selection.TypeText(%s)", @text)


  ;Find & Replace in Word
  dhCallMethod    (oWord,"Selection.Goto = %u",0,0,0)

  dhPutValue      (oWord,"Selection.Find.Forward = %b",#True)
  dhPutValue      (oWord,"Selection.Find.MatchWholeWord = %b",#True)
  dhPutValue      (oWord,"Selection.Find.Text = %s",@"Sample")
  dhPutValue      (oWord,"Selection.Find.Replacement.Text = %s",@"DONE")

  dhCallMethod    (oWord, "Selection.Find.Execute(%m,%m,%m,%m,%m,%m,%m,%m,%m,%m,%u)",wdReplaceAll)

  dhReleaseObject (oWord) : oWord = 0

EndIf
