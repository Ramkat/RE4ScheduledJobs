'Filename    : AnalyseFTPLogs.vbs
'Author      : Christo Pretorius
'Date        : 30 June 2004
'Description : This script will run the cmd line to analyse the FT logs.
'		All files of 2 months ago will be deleted.

Option Explicit

CONST conDebugMode = False

Dim objFSO
Dim objShell
Dim strLogFiles
Dim strDelFiles
Dim dteDate
Dim strDelDate
Dim strDate
Dim intYear
Dim intMonth
Dim strCmd
Dim lngReturnValue
Dim strLogFilesPath

strLogFilesPath = "D:\IIS Logs\FTP\MSFTPSVC1"

'Create an object of the cmd shell
Set objShell = CreateObject("WScript.Shell")

'Create an object of the file system
Set objFSO = CreateObject("Scripting.FileSystemObject")

dteDate = Date()
'For testing dteDate = CDate("2003/11/01")	'###
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
	'this month - 2 month's files. E.g. If it is Feb 2004 you want to delete Dec 2003's files.
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

'Analyse the logs
strCmd = "D:\LogAnalysers\IIS_5_FTP_Log_Analyser.exe " & strLogFilesPath & "\" & strLogFiles & ",D:\Webs\FTPStats"
											
On Error Resume Next		
If conDebugMode Then
	wscript.echo "Command = " & vbcrlf & strCmd
Else
	lngReturnValue = objShell.Run(strCmd, 7, True)				
End If
On Error Goto 0
			
If conDebugMode Then
	wscript.echo "Deleting files: " & strLogFilesPath & "\" & strDelFiles
Else				
	'Delete the old files
	On Error Resume Next
	objFSO.DeleteFile strLogFilesPath & "\" & strDelFiles, True				
	On Error Goto 0
	Err.Clear
End If

Set objFSO = Nothing
Set objShell = Nothing

If conDebugMode Then
	wscript.echo "Done"
End If