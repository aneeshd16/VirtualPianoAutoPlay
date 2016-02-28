'Time interval between notes'
timeInterval = 100

Dim InputFile 
Dim FSO, oFile 
Dim strData 
InputFile = "sheet.txt" 
Set FSO = CreateObject("Scripting.FileSystemObject") 
Set oFile = FSO.OpenTextFile(InputFile) 
str = oFile.ReadAll 
oFile.Close 

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.AppActivate("chrome")
WScript.Sleep 2000
interval = timeInterval
For i=1 To Len(str)
	If Mid(str,i,1) = "[" Then
		interval = 0
	ElseIf Mid(str,i,1) = "]" Then
		interval = timeInterval
		WScript.Sleep interval
	Else
		WshShell.SendKeys Mid(str,i,1)
		WScript.Sleep interval
	End If
Next