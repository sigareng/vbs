dim NIC1, Nic, StrIP, CompName

Set NIC1 = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")

For Each Nic in NIC1

if Nic.IPEnabled then

StrIP = Nic.IPAddress(i)

Set WshNetwork = WScript.CreateObject("WScript.Network")

CompName= WshNetwork.Computername

MsgBox "IP Address:  "&StrIP & vbNewLine _
& "Computer Name:  "&CompName,4160,"IP Address and Computer Name"

wscript.quit

end if

next