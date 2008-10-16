?; by ts-soft

Define.l Result, obj = dhCreateObject("WScript.Network")
Define.s szUserDomain, szComputerName, szUserName, Message.s

dhToggleExceptions(#True)

If obj

  dhGetValue("%T", @Result, obj, "UserDomain")
  If Result
    szUserDomain = PeekS(Result)
    dhFreeString(Result) : Result = 0
  EndIf

  dhGetValue("%T", @Result, obj, "ComputerName")
  If Result
    szComputerName = PeekS(Result)
    dhFreeString(Result) : Result = 0
  EndIf

  dhGetValue("%T", @Result, obj, "UserName")
  If Result
    szUserName = PeekS(Result)
    dhFreeString(Result) : Result = 0
  EndIf

  dhReleaseObject(obj) : obj = 0

EndIf

Message = "UserDomain:"   + #TAB$ + szUserDomain    + #LF$
Message + "ComputerName:" + #TAB$ + szComputerName  + #LF$
Message + "UserName:"     + #TAB$ + szUserName

MessageRequester("Network-Info", Message, #MB_ICONINFORMATION)
