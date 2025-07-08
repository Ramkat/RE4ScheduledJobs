'Filename    : AnalyseIISLogs-Daily.vbs
'Author      : Christo Pretorius
'Date        : 11 Feb 2009
'Description : This script will traverse through the iis logs folder,
'	 	copy the analog-daily.cfg file to c:\analog\analog.cfg, and then
'		run the cmd line to analyse the logs.

Option Explicit

CONST conDebugMode = False

Dim objFSO
Dim objFile
Dim objShell
Dim objIISFolder
Dim objWebFolders
Dim objSubFolders
Dim WebFolder
Dim SubFolder
Dim strLogFiles
Dim dteDate
Dim strDate
Dim strAnalysedFile
Dim strAnalysedFolder
Dim intYear
Dim intMonth
Dim strCmd
Dim lngReturnValue
Dim sinTotalKiloBytes
Dim strDateToday
Dim strWebsiteValue

'Create an object of the cmd shell
Set objShell = CreateObject("WScript.Shell")

'Create an object of the file system
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Create an object of the IIS Logs folder
Set objIISFolder = objFSO.GetFolder("c:\IIS Logs")

'Create an object for all the web folders of the "IISLogs" folder
Set objWebFolders = objIISFolder.SubFolders

dteDate = Date()
'for testing dteDate = CDate("2003/11/01")	'###
intYear = Year(dteDate)	'Get the year
intMonth = Month(dteDate)
strDate = Trim(CStr(intYear))		'Convert the year to a string	
strAnalysedFile = "Stats" & strDate
strDateToday = MonthName(intMonth) & " " & strDate
strDate = Mid(strDate, 3)				'Remove the century from the string	
	
'If the month is less than 10, then add a "0" in front of it.
If intMonth < 10 Then
	strDate = strDate & "0" & Trim(CStr(intMonth))
	strAnalysedFile = strAnalysedFile & "0" & Trim(CStr(intMonth))	
Else
	strDate = strDate & Trim(CStr(intMonth))
	strAnalysedFile = strAnalysedFile & Trim(CStr(intMonth))
End If

strAnalysedFile = strAnalysedFile & ".html"
sinTotalKiloBytes = 0
strWebsiteValue = ""
	
If conDebugMode Then
	wscript.echo "strDate = " & strDate
End If

'Logfile names are in the format: exMMYYDD.log.
'We want to analyse per month e.g. ex0930*.log
strLogFiles = "ex" & strDate & "*.log"

If conDebugMode Then
	wscript.echo "strLogFiles = " & strLogFiles
End If

'Loop through all the web folders
For Each WebFolder In objWebFolders ' c:\iis logs\wildlifecampus etc.
	'If the Analog file exists, copy it to the Analog folder, 
	'delete the old files and analyse the latest ones.		
	If objFSO.FileExists(objIISFolder.Path & "\" & WebFolder.Name & "\analog-daily.cfg") Then
		If conDebugMode Then
			wscript.echo "Copying " & objIISFolder.Path & "\" & WebFolder.Name & "\analog-daily.cfg" & vbcrlf & "to " & "c:\analog\analog.cfg"
		Else	
			'Copy the analog.cfg file for the specific website to the Analog folder
			objFSO.CopyFile objIISFolder.Path & "\" & WebFolder.Name & "\analog-daily.cfg", "c:\analog\analog.cfg", True
			strAnalysedFolder = "c:\Webs\AnalogStats\" & WebFolder.Name
			strWebsiteValue = strWebsiteValue & "<br><b>" & WebFolder.Name & "</b>"
		End If
				
		'Create an object of the subfolder's subfolders
		Set objSubFolders = WebFolder.SubFolders
	
		'Loop through all the sub folders of the web folder (should only be W3SVC1 etc.)
		For Each SubFolder In objSubFolders ' c:\iis logs\wildlifecampus\w3svc1 etc.
			If conDebugMode Then
				wscript.echo "The command is:" & vbcrlf & "c:\analog\analog.exe " & chr(34) & SubFolder.Path & "\" & strLogFiles & chr(34)
			Else
				'Analyse the logs
				strCmd = "c:\analog\analog.exe " & chr(34) & SubFolder.Path & "\" & strLogFiles & chr(34)
											
				On Error Resume Next		
				lngReturnValue = objShell.Run(strCmd, 7, True)				
				On Error Goto 0			 								
				
				'Sum the report's usage data.
				Call SumTotalBytes(strAnalysedFolder & "\" & strAnalysedFile)
			End If
		Next
	End If	
Next

Set objSubFolders = Nothing
Set objWebFolders = Nothing
Set objIISFolder = Nothing

Call WriteSummary
Set objFSO = Nothing
Set objShell = Nothing

If conDebugMode Then
	wscript.echo "Done"
End If

Sub SumTotalBytes(strPathAndFile)
	'This sub will open the newly created report file, search for the "Data transferred:" line,
	'and add the value of the transferred data to the global variable sinTotalKiloBytes.
	
	'Check if the file exists.
	If objFSO.FileExists(strPathAndFile) = False Then Exit Sub
			
	Dim intPos	
	Dim strLine		
	Dim sinValue
	Dim sinKiloBytes
	
	'Open the file for reading.
	Set objFile = objFSO.OpenTextFile(strPathAndFile, 1, False)
	
	'Read the 1st line.
'###	strLine = LCase(objFile.ReadLine)

	
	'Loop through the file until the line is found that contains the value we need.
	Do While Not objFile.AtEndOfStream
		strLine = LCase(objFile.ReadLine)

	
	'If the string is found...
	If Left(strLine, 28) = "<br><b>data transferred:</b>" Then	
		'Find the 1st "("		
		intPos = InStr(1, strLine , "(") 		
		
		If intPos > 0 Then 			
			'Get the text that holds the value.
			strLine = Trim(Mid(strLine, 29, intPos - 29))
			
			strWebsiteValue = strWebsiteValue & " " & strLine
			
			'Extract the numeric value.
			intPos = InStr(1, strLine, " ")
			sinValue = Trim(Left(strLine, intPos - 1))
			sinValue = CSng(sinValue)
			
			'Now determine the size of the transfered data.
			If InStr(1, strLine, "gigabytes") > 0 Then
				sinKiloBytes = sinValue * 1024 'Convert to megabytes
				sinKiloBytes = sinValue * 1024 'Convert to kilobytes
			End If
						
			If InStr(1, strLine, "megabytes") > 0 Then				
				sinKiloBytes = sinValue * 1024 'Convert to kilobytes
			End If
						
			If InStr(1, strLine, "kilobytes") > 0 Then				
				sinKiloBytes = sinValue
			End If
				
			'Add the value to the global variable.
			sinTotalKiloBytes = sinTotalKiloBytes + sinKiloBytes
		End If
	End If
		
	Loop
	objFile.Close
	Set objFile = Nothing
	Exit Sub
End Sub

Sub WriteSummary
	'This sub will create a summary of all the stats in the RE4 folder.
	
	Dim strHTML
	Dim strValue
	Dim strFile
			
	strValue = "Kilobytes"
	
	If sinTotalKiloBytes > 1024 Then
		sinTotalKiloBytes = sinTotalKiloBytes / 1024	'Convert to megabytes
		strValue = "Megabytes"
	End If
	
	If sinTotalKiloBytes > 1024 Then
		sinTotalKiloBytes = sinTotalKiloBytes / 1024	'Convert to gigabytes
		strValue = "Gigabytes"
	End If
	
	strHTML = "<html><head><title>Summed website stats for " & strDateToday & "</title><body><h3>" _
		& "Summed website stats for " & strDateToday & "</h3><br>" & strWebsiteValue _
		& "<br><br><b>Total data transferred:</b> " & FormatNumber(sinTotalKiloBytes, 2,,,True) _
		& " " & strValue & "</head></html>"
		
	'Create file, write to it, close it.
	On Error Resume Next
	strFile = "c:\webs\AnalogStats\re4.co.za\Summed" & strAnalysedFile
	
	If objFSO.FileExists(strFile) Then objFSO.DeleteFile(strFile)
	
	Set objFile = objFSO.OpenTextFile(strFile, 8, True)
	objFile.WriteLine strHTML
	objFile.Close
	Set objFile = Nothing	
End Sub