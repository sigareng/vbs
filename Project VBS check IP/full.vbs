strAddress = InputBox ("Specify an address to ping", WScript.ScriptFullName, "google.com")

If GetOsVersion < "5.01" Then
  MsgBox "Unsupported Operating System", _
         vbOKOnly  + vbExclamation, WScript.ScriptFullName
  WScript.Quit
End If

Set objPing = GetObject("winmgmts:").Get("Win32_PingStatus.Address='" & strAddress & "'")

With objPing
  Wscript.Echo "Address : " & .Address
  Wscript.Echo "Buffer size : " & .BufferSize
  Wscript.Echo "No Fragmentation : " & .NoFragmentation
  Wscript.Echo "PrimaryAddressResolutionStatus : " & .PrimaryAddressResolutionStatus
  Wscript.Echo "ProtocolAddress : " & .ProtocolAddress
  Wscript.Echo "ProtocolAddressResolved : " & .ProtocolAddressResolved
  Wscript.Echo "RecordRoute : " & .RecordRoute
  Wscript.Echo "ReplyInconsistency : " & .ReplyInconsistency
  Wscript.Echo "ReplySize : " & .ReplySize
  Wscript.Echo "ResolveAddressNames : " & .ResolveAddressNames
  Wscript.Echo "ResponseTime : " & .ResponseTime
  Wscript.Echo "ResponseTimeToLive : " & .ResponseTimeToLive
  If IsNull (.RouteRecord) Then
    Wscript.Echo "RouteRecord : Null"
  Else
    Wscript.Echo "RouteRecord : " & _
           Join (.RouteRecord, "; ")
  End If
  If IsNull (.RouteRecordResolved) Then
    Wscript.Echo "RouteRecordResolved : Null"
  Else
    Wscript.Echo "RouteRecordResolved : " & _
           Join (.RouteRecordResolved, "; ")
  End If
  Wscript.Echo "SourceRoute : " & .SourceRoute
  Wscript.Echo "SourceRouteType : " & GetSourceRouteType(.SourceRouteType)
  Wscript.Echo "Status code : " & GetStatusCode(.StatusCode)
  Wscript.Echo "Timeout : " & .TimeOut
  If IsNull (.TimeStampRecord) Then
    Wscript.Echo "TimeStampRecord : Null"
  Else
    Wscript.Echo "TimeStampRecord : " & _
           Join (.TimeStampRecord, "; ")
  End If
  If IsNull (.TimeStampRecordAddress) Then
    Wscript.Echo "TimeStampRecordAddress : Null"
  Else
    Wscript.Echo "TimeStampRecordAddress : " & _
           Join (.TimeStampRecordAddress, "; ")
  End If
  If IsNull (.TimeStampRecordAddressResolved) Then
    Wscript.Echo "TimeStampRecordAddressResolved : Null"
  Else
    Wscript.Echo "TimeStampRecordAddressResolved : " & _
           Join (.TimeStampRecordAddressResolved, "; ")
  End If
  Wscript.Echo "TimeStampRoute : " & .TimeStampRoute
  Wscript.Echo "TimeToLive : " & .TimeToLive
  Wscript.Echo "TypeOfService : " & GetTypeOfService(.TypeOfService)
End With

MsgBox "The END"

' ___________________
Function GetOsVersion
  Set objWMIService = GetObject("winmgmts:") 
  Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_OperatingSystem",,48) 
  For Each objItem In colItems
    WScript.Echo objItem.Version
    GetOsVersion = objItem.Version
  Next
  Set colItems = Nothing : Set objWMIService = Nothing
End Function

' ____________________________________________
Function GetSourceRouteType (intSourceRouteType)

  Dim strType

  Select Case intSourceRouteType
  case 1
    strType = "Loose Source Routing"
  case 2
    strType = "Strict Source Routing"
  case Else
    ' Default - 0 - or any other value.
    strType = intSourceRouteType & " - None"
  End Select
  GetSourceRouteType = strType

End Function

' ______________________________________
Function GetTypeOfService (intServiceType)

  Dim strType

  Select Case intServiceType
  case 2
    strType = "Minimize Monetary Cost"
  case 4
    strType = "Maximize Reliability"
  case 8
    strType = "Maximize Throughput"
  case 16
    strType = "Minimize Delay"
  Case Else
    ' Default - 0 - or any other value.
    strType = intServiceType & " - Normal"
  End Select
  GetTypeOfService = strType

End Function

' ____________________________
Function GetStatusCode (intCode)

  Dim strStatus
  Select Case intCode
  case  0
    strStatus = "Success"
  case  11001
    strStatus = "Buffer Too Small"
  case  11002
    strStatus = "Destination Net Unreachable"
  case  11003
    strStatus = "Destination Host Unreachable"
  case  11004
    strStatus = "Destination Protocol Unreachable"
  case  11005
    strStatus = "Destination Port Unreachable"
  case  11006
    strStatus = "No Resources"
  case  11007
    strStatus = "Bad Option"
  case  11008
    strStatus = "Hardware Error"
  case  11009
    strStatus = "Packet Too Big"
  case  11010
    strStatus = "Request Timed Out"
  case  11011
    strStatus = "Bad Request"
  case  11012
    strStatus = "Bad Route"
  case  11013
    strStatus = "TimeToLive Expired Transit"
  case  11014
    strStatus = "TimeToLive Expired Reassembly"
  case  11015
    strStatus = "Parameter Problem"
  case  11016
    strStatus = "Source Quench"
  case  11017
    strStatus = "Option Too Big"
  case  11018
    strStatus = "Bad Destination"
  case  11032
    strStatus = "Negotiating IPSEC"
  case  11050
    strStatus = "General Failure"
  case Else
    strStatus = intCode & " - Unknown"
  End Select
  GetStatusCode = strStatus

End Function