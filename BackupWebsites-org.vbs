'Filename    : BackupWebsites.vbs
'Author      : Christo Pretorius
'Date        : 9 September 2003
'Description : This script will copy the website folders to the db server.
'            : It will delete backup files and folders older than conBackupDays 
'	       and copy the latest files.

Option Explicit

CONST conBackupFolder = "C:\webs"
CONST conCopyToFolder = "D:\Webs-Backup"
CONST conBackupDays = 14
CONST conDebugMode = False

Dim objFSO
Dim objFolder
Dim objSubFolders
Dim objWebFolders
Dim objFile
Dim objTS
Dim objShell
Dim SubFolder
Dim WebFolder
Dim dteDate
Dim strDate
Dim strTestDate
Dim strReport

'Create an object of the file system
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Create an object of the "Copy To" folder
Set objFolder = objFSO.GetFolder(conCopyToFolder)

'Create an object for all the sub folders of the "Copy To" folder
Set objSubFolders = objFolder.SubFolders

'Create an object of the cmd shell
Set objShell = CreateObject("WScript.Shell")

dteDate = Date()
strDate = strGetDateAsNumerics(dteDate)
strReport = "<h4>Website backups on " & Trim(CStr(Year(dteDate))) & "/" & Trim(CStr(Month(dteDate))) & "/" & Trim(CStr(Day(dteDate))) & " at " & Time() & "</h4>" & vbcrlf

dteDate = Date() - conBackupDays
strTestDate = strGetDateAsNumerics(dteDate)

'=====
'Delete all the backup folders older than conBackupDays in the "Copy To" folder
'and create a new folder with today's date for the new backup files.
'=====
'Loop through all the sub folders
For Each SubFolder In objSubFolders

	'Create an object of the subfolder's subfolders
	Set objWebFolders = SubFolder.SubFolders
	
	'Loop through all the sub folders of the web folder
	For Each WebFolder In objWebFolders
	
		'If the web sub folder is older than conBackupDays or its name is older than the date minus conBackupDays, delete it.
		If WebFolder.DateCreated < dteDate Or WebFolder.Name = strTestDate Then
			If Not conDebugMode Then
				On Error Resume Next
				objFSO.DeleteFolder conCopyToFolder & "\" & SubFolder.Name & "\" & WebFolder.Name, True
				On Error Goto 0
			Else
				wscript.echo "Deleting " & conCopyToFolder & "\" & SubFolder.Name & "\" & WebFolder.Name
			End If		
		End If
	Next
	
	Set objWebFolders = Nothing
Next
	
Set objSubFolders = Nothing	
Set objFolder = Nothing

'=====
'Copy the website folders to the conCopyToFolder.
'=====
'Create an object of the backups folder
Set objFolder = objFSO.GetFolder(conBackupFolder)

'Create an object for all the sub folders of the backups folder
Set objSubFolders = objFolder.SubFolders

'Loop through all the sub folders
For Each SubFolder In objSubFolders
	'Check if the target folder exists.
	strReport = strReport & "<ul><li>" & SubFolder.Name
	
	If Not objFSO.FolderExists(conCopyToFolder & "\" & SubFolder.Name) Then
		If Not conDebugMode Then							
			On Error Resume Next
			objFSO.CreateFolder(conCopyToFolder & "\" & SubFolder.Name)
			On Error Goto 0
		Else
			wscript.echo "Creating folder " & conCopyToFolder & "\" & SubFolder.Name
		End If
	End If
					
	'Check if the target folder with today's date exists.
	strReport = strReport & "<ul><li>" & strDate & "</li></ul>" & vbcrlf
	If Not objFSO.FolderExists(conCopyToFolder & "\" & SubFolder.Name & "\" & strDate) Then
		If Not conDebugMode Then
			On Error Resume Next
			objFSO.CreateFolder(conCopyToFolder & "\" & SubFolder.Name & "\" & strDate)
			On Error Goto 0
		Else
			wscript.echo "Creating folder " & conCopyToFolder & "\" & SubFolder.Name & "\" & strDate
		End If
	End If
															
	If Not conDebugMode Then
		On Error Resume Next		
		lngReturnValue = objShell.Run("robocopy.exe " & conBackupFolder & "\" & SubFolder.Name & " " & conCopyToFolder & "\" & SubFolder.Name & "\" & strDate & " /E /R:10 /W:30 /NP /XD Upload", 7, True)
		On Error Goto 0
	Else
		wscript.echo "Copying folder " & conBackupFolder & "\" & SubFolder.Name & vbcrlf & "to " & conCopyToFolder & "\" & SubFolder.Name & "\" & strDate & vbcrlf _
			& "robocopy.exe " & conBackupFolder & "\" & SubFolder.Name & " " & conCopyToFolder & "\" & SubFolder.Name & "\" & strDate & " /E /R:10 /W:30 /NP /XD Upload"
	End If
	
	strReport = strReport & "</li></ul>" & vbcrlf
Next			

Set objSubFolders = Nothing
Set objFolder = Nothing

'The BackupWebsitesToFTP.vbs script doesn't want to work from here. 
'It now scheduled via Windows Scheduler.
'Start the script to copy from the website BACKUP folder to the FTP folder.
'On Error Resume Next	
'strReport = strReport & "<ul><li>FTP backup script date: " & strDate & ". Return value: "	
'lngReturnValue = objShell.Run("wscript.exe C:\ScheduledJobs\BackupWebsitesToFTP.vbs " & strDate, 7, False)
'strReport = strReport & lngReturnValue &".</li></ul>"
'On Error Goto 0

strReport = "<ul<li>Website backups ended at " & Time() & ".</li></ul>"

'Overwrite the current backup report
objFSO.CreateTextFile conBackupFolder & "\re4.co.za\WebBackups.htm", True
Set objFile = objFSO.GetFile(conBackupFolder & "\re4.co.za\WebBackups.htm")
Set objTS = objFile.OpenAsTextStream(2, 0)
objTS.Write strReport
objTS.Close

Set objTS = Nothing
Set objFile = Nothing
Set objFSO = Nothing
Set objShell = Nothing

If conDebugMode Then	wscript.echo "Done"
Wscript.Quit

Function strGetDateAsNumerics(dteDate)
	'This function will return the date in the format yyyymmdd as a string.
	'dteDate = The date to convert in datetime format.
	
	strGetDateAsNumerics = ""			'Assume failure
	
	'Test if the parameter is a valid date.
	If Not IsDate(dteDate) Then Exit Function
	
	Dim strYear
	Dim strMonth
	Dim strDay		
	
	strYear = Year(dteDate)			'Get the year
	strMonth = Month(dteDate)		'Get the month
	strDay = Day(dteDate)				'Get the day
	
	'Ensure that the month is 2 digits.
	If strMonth < 10 Then
		strMonth = "0" & strMonth
	End If
	
	'Ensure that the day is 2 digits.
	If strDay < 10 Then
		strDay = "0" & strDay
	End If	
	
	'Return the value.
	strGetDateAsNumerics = strYear & strMonth & strDay
End Function