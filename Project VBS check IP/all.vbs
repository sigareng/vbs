VBSTART
Function Ping(strHost)
  dim objPing, objRetStatus

  ipAddress = ""
  set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery ("select * from Win32_PingStatus where address = '" & strHost & "'")

  for each objRetStatus in objPing
    if IsNull(objRetStatus.StatusCode) or objRetStatus.StatusCode<>0 then
      Ping = ""
    else
      Ping = objRetStatus.ProtocolAddress
    end if
  next
End Function
VBEND

VBEval>Ping("http://officialsaipul.cf"),pinged
If>pinged<>{""}
  MessageModal>Ip Address is: %pinged%
Else
  MessageModal>host not contacted
Endif