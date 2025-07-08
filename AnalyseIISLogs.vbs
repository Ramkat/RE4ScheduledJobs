'Filename    : AnalyseIISLogs.vbs
'Author      : Christo Pretorius
'Date        : 17 October 2003
'Description : This script will traverse through the iis logs folder,
'	 	copy the analog.cfg file to c:\analog, and then run
'		the cmd line to analyse the log.
'		All files of 2 months ago will be deleted.

Option Explicit

CONST conDebugMode = False

Dim objFSO
Dim objShell
Dim objIISFolder
Dim objWebFolders
Dim objSubFolders
Dim WebFolder
Dim SubFolder
Dim strLogFiles
Dim strDelFiles
Dim dteDate
Dim strDelDate
Dim strDate
Dim intYear
Dim intMonth
Dim strCmd
Dim lngReturnValue

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

If conDebugMode Then
	wscript.echo "intYear/intMonth = " & intYear & "/" & intMonth
End If

'If it is the 1st month, then use last year and December.
If intMonth = 1 Then
	intYear = intYear - 1
	strDate = Trim(CStr(intYear))		'Convert the year to a string	
	strDate = Mid(strDate, 3)				'Remove the century from the string	
	strDelDate = strDate & "11"
	strDate = strDate & "12"
Else
	intMonth = intMonth - 1					'Subtract 1 month because we want to analyse last month's logs
	strDate = Trim(CStr(intYear))		'Convert the year to a string
	strDate = Mid(strDate, 3)				'Remove the century from the string

	'If the month is less than 10, then add a "0" in front of it.
	If intMonth < 10 Then
		strDate = strDate & "0" & Trim(CStr(intMonth))
	Else
		strDate = strDate & Trim(CStr(intMonth))
	End If
	
	'If it is month 1, delete 1 from the year because you want to delete
	'this month - 2 month's files. E.g. If it is Feb 2003 you want to delete Dec 2003's files.
	If intMonth = 1 Then
		intYear = intYear - 1
		strDelDate = Trim(CStr(intYear))		'Convert the year to a string
		strDelDate = Mid(strDelDate, 3)				'Remove the century from the string	
		strDelDate = strDelDate & "12"
	Else
		intMonth = intMonth - 1			'Subtract another 1 month because we want to delete the files of 2 monts ago.
		strDelDate = Trim(CStr(intYear))
		strDelDate = Mid(strDelDate, 3)
		
		'If the month is less than 10, then add a "0" in front of it.
		If intMonth < 10 Then
			strDelDate = strDelDate & "0" & Trim(CStr(intMonth))
		Else
			strDelDate = strDelDate & Trim(CStr(intMonth))
		End If
	End If				
End If

'Logfile names are in the format: exMMYYDD.log.
'We want to analyse per month e.g. ex0930*.log
strLogFiles = "ex" & strDate & "*.log"
strDelFiles = "ex" & strDelDate & "*.log"

If conDebugMode Then
	wscript.echo "strLogFiles = " & strLogFiles & vbcrlf & "strDelFiles = " & strDelFiles		
End If

'Loop through all the web folders
For Each WebFolder In objWebFolders ' c:\iis logs\wildlifecampus etc.
	'If the Analog file exists, copy it to the Analog folder, 
	'delete the old files and analyse the latest ones.		
	If objFSO.FileExists(objIISFolder.Path & "\" & WebFolder.Name & "\analog.cfg") Then
		'Copy the analog.cfg file for the specific website to the Analog folder
		objFSO.CopyFile objIISFolder.Path & "\" & WebFolder.Name & "\analog.cfg", "c:\analog\analog.cfg", True
				
		'Create an object of the subfolder's subfolders
		Set objSubFolders = WebFolder.SubFolders
	
		'Loop through all the sub folders of the web folder (should only be W3SVC1 etc.)
		For Each SubFolder In objSubFolders ' c:\iis logs\wildlifecampus\w3svc1 etc.
			'Analyse the logs
			strCmd = "c:\analog\analog.exe " & chr(34) & SubFolder.Path & "\" & strLogFiles & chr(34)
											
			On Error Resume Next		
			lngReturnValue = objShell.Run(strCmd, 7, True)
If conDebugMode Then
	If Err.Number <> 0 Then
		wscript.echo "Err = " & Err.Description	
	End If
End If
			On Error Goto 0			 								

			
			'Delete the old files
			On Error Resume Next
			objFSO.DeleteFile SubFolder.Path & "\" & strDelFiles, True
				
			'If Err.Number > 0 Then
			'	If Err.Number = 53 Then		'53 = file not found - ignore it
			'		'Do nothing
			'	Else
			'		wscript.end
			'	End If
			'End If
			
			On Error Goto 0
			Err.Clear

		Next
	End If	
Next

Set objSubFolders = Nothing
Set objWebFolders = Nothing
Set objIISFolder = Nothing
Set objFSO = Nothing
Set objShell = Nothing

If conDebugMode Then
	wscript.echo "Done"
End If