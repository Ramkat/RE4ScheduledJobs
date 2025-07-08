'Filename	: Reboot.vbs
'Author		: Christo Pretorius
'Date		: 21 October 2008
'Description:	Reboots the server and logs the date & time of the reboot.

wscript.quit 0	'Added to prevent the script from executing. Job also disabled in robotask.

Dim objFSO
Dim objFile
Dim objShell
Dim lngReturnValue
Dim dteNow

dteNow = Now()

'Create an object of the file system.
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("Reboot.log", 2, True)
'Create an object of the cmd shell.
Set objShell = CreateObject("WScript.Shell")

'Update the reboot log.
objFile.Write "Reboot initiated on " & GetWeekday & " " & Day(dteNow) _
	& " " & MonthName(Month(dteNow)) & " " & Year(dteNow) & " at " & Hour(dteNow) _
	& ":" & Minute(dteNow) & ":" & Second(dteNow)
	
objFile.Close
Set objFile = Nothing
Set objFSO = Nothing	

'Call the command that will reboot the server.
lngReturnValue = objShell.Run("c:\windows\system32\shutdown.exe /r /f /t 30", 7, False)
wscript.quit 0

Function GetWeekday
	Select Case Weekday(dteNow)
		Case vbSunday			GetWeekday = "Sunday"
		Case vbMonday			GetWeekday = "Monday"
		Case vbTuesday		GetWeekday = "Tusday"
		Case vbWednesday	GetWeekday = "Wednesday"
		Case vbThursday		GetWeekday = "Thursday"
		Case vbFriday			GetWeekday = "Friday"
		Case vbSaturday		GetWeekday = "Saturday"
	End Select
End Function