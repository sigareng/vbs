Dim string 
String = "This Is A script that allows me to copy text contained withing these quotes directly into my clipboard. Which yes is plenty fast as it is when compared to finding file, opening file, selecting desired content, copy content, and select location to paste content."
Set WshShell = WScript.CreateObject("WScript.Shell") 
WshShell.Run "cmd.exe /c echo " & String & " | clip", 0, TRUE