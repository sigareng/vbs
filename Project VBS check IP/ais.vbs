VBSTART
Function Ping(strHost)
  dim objPing, objRetStatus

  set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery ("select * from Win32_PingStatus where address = '" & strHost & "'")

  for each objRetStatus in objPing
    if IsNull(objRetStatus.StatusCode) or objRetStatus.StatusCode<>0 then
      Ping = False
      'MsgBox "Status code is " & objRetStatus.StatusCode
    else
      Ping = True
      'MsgBox "Bytes = " & vbTab & objRetStatus.BufferSize & vbCRLF & _
      '       "Time (ms) = " & vbTab & objRetStatus.ResponseTime & vbCRLF & _
      '       "TTL (s) = " & vbTab & objRetStatus.ResponseTimeToLive
    end if
  next
End Function
VBEND


Label>PingLoop
//Wait 1 second
wait>1
//Change to any host IP you like. I'm using Google's IP in this example.
VBEval>Ping("72.14.209.104"),pinged
If>pinged=True
  MessageModal>host contacted
Else
  MessageModal>host not contacted
  //Restart Modem/PC
Endif
Goto>PingLoop