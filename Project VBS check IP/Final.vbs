strAddress = InputBox ("Masukan Alamat URL dengan benar", WScript.ScriptFullName, "officialsaipul.cf")

Set objPing = GetObject("winmgmts:").Get("Win32_PingStatus.Address='" & strAddress & "'")

With objPing
  Wscript.Echo "Alamat : " & .Address
  Wscript.Echo "IP Terbaca : " & .ProtocolAddress
  Dim string 
String = ""& .ProtocolAddress
Set WshShell = WScript.CreateObject("WScript.Shell") 
WshShell.Run "cmd.exe /c echo " & String & " | clip", 0, TRUE
End With

MsgBox "Selesai"