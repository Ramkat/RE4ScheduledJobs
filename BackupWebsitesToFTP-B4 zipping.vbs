'Filename    : BackupWebsitesToFTP.vbs
'Author      : Christo Pretorius
'Date        : 27 October 2010
'Description : This script will copy the website backup from the supplied folder date to the FTP server.
'            : It will delete backup files and folders older than conBackupDays .
'Parameter   : Folder date in the format yyyymmdd

Option Explicit

'DISABLED wscript.quit 'Exit without execution

CONST conBackupFolder = "C:\Webs-Backup"
CONST conCopyToFolder = "D:"
CONST conBackupDays = 7
CONST conDebugMode = false

Dim strFolderDate
Dim strLongDate
Dim objFSO
Dim objFolder
Dim objSubFolders
Dim objWebFolders
Dim objTS
Dim objShell
Dim SubFolder
Dim WebFolder
Dim dteDate
Dim strDate
Dim strTestDate
Dim intDay

'If no argument is passed, quit.
If WScript.Arguments.Count <> 1 Then Wscript.Quit
	
strFolderDate = WScript.Arguments.Item(0) 'Argument passed to the script file. It is the folder date for the backups in format yyyymmdd.

strLongDate = Left(strFolderDate, 4) & "/" & Mid(strFolderDate, 5, 2) & "/" & Mid(strFolderDate, 7, 2)

If conDebugMode Then Wscript.echo "strLongDate = " & strLongDate

'Don't copy the backups on a Saturday, Sunday or Monday morning due to server problems.
intDay = Weekday(CDate(strLongDate))

If conDebugMode Then Wscript.echo "intDay = " & intDay

If intDay = vbSaturday or intDay = vbSunday or intDay = vbMonday Then
	Wscript.Quit
End If

'Create an object of the file system
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Create an object of the "Copy To" folder
Set objFolder = objFSO.GetFolder(conCopyToFolder)

'Create an object for all the sub folders of the "Copy To" folder
Set objSubFolders = objFolder.SubFolders

'Create an object of the cmd shell
Set objShell = CreateObject("WScript.Shell")

'Get the date less conBackupDays.
dteDate = CDate(strLongDate) - conBackupDays

If conDebugMode Then Wscript.echo "dteDate = " & dteDate

strTestDate = strGetDateAsNumerics(dteDate)

If conDebugMode Then Wscript.echo "strTestDate = " & strTestDate

'=====
'Delete all the backup folders older than conBackupDays in the "Copy To" folder.
'=====
'Loop through all the sub folders
For Each SubFolder In objSubFolders
	'If the web sub folder is older than conBackupDays or its name is older than the date minus conBackupDays, delete it.
	If SubFolder.DateCreated < dteDate Or SubFolder.Name = strTestDate Then
		If Not conDebugMode Then
			On Error Resume Next
			objFSO.DeleteFolder conCopyToFolder & "\" & SubFolder.Name, True
			On Error Goto 0
		Else
			wscript.echo "Deleting " & conCopyToFolder & "\" & SubFolder.Name
		End If		
	End If
Next
	
Set objSubFolders = Nothing	
Set objFolder = Nothing

'=====
'Copy the website folders to the conCopyToFolder.
'=====
'Create an object of the backups folder
Set objFolder = objFSO.GetFolder(conBackupFolder)

'Create an object for all the sub folders of the backups folder (the website names)
Set objSubFolders = objFolder.SubFolders

'Check if the target folder with the folder date passed as an argument exists.
If Not objFSO.FolderExists(conCopyToFolder & "\" & strFolderDate) Then
	If Not conDebugMode Then
		On Error Resume Next
		objFSO.CreateFolder(conCopyToFolder & "\" & strFolderDate)
		On Error Goto 0
	Else
		wscript.echo "Creating folder " & conCopyToFolder & "\" & strFolderDate
	End If
End If

'Loop through all the sub folders of the backups folder (the website names)
For Each SubFolder In objSubFolders
	'Create the target (websitename) folder.
	If Not conDebugMode Then							
		On Error Resume Next
		objFSO.CreateFolder(conCopyToFolder & "\" & strFolderDate & "\" & SubFolder.Name)
		On Error Goto 0
	Else
		wscript.echo "Creating folder " & conCopyToFolder & "\" & strFolderDate & "\" & SubFolder.Name
	End If
															
	If Not conDebugMode Then
		On Error Resume Next
		lngReturnValue = objShell.Run("robocopy.exe " & conBackupFolder & "\" & SubFolder.Name & "\" & strFolderDate & " " & conCopyToFolder & "\" & strFolderDate & "\" & SubFolder.Name & " /E /R:10 /W:30 /NP", 7, True)
		On Error Goto 0
	Else
		wscript.echo "Copying folder " & conBackupFolder & "\" & SubFolder.Name & "\" & strFolderDate & vbcrlf & "to " & conCopyToFolder & "\" & strFolderDate & "\" & SubFolder.Name & vbcrlf _
			& "robocopy.exe " & conBackupFolder & "\" & SubFolder.Name & "\" & strFolderDate & " " & conCopyToFolder & "\" & strFolderDate & "\" & SubFolder.Name & " /E /R:10 /W:30 /NP"
	End If
Next

'Backup a few ad-hoc folders
If Not conDebugMode Then
  On Error Resume Next
  lngReturnValue = objShell.Run("robocopy.exe C:\MiscInfo " & conCopyToFolder & "\" & strFolderDate & "\AdHoc\MiscInfo /E /R:10 /W:30 /NP", 7, True)
  lngReturnValue = objShell.Run("robocopy.exe C:\ScheduledJobs " & conCopyToFolder & "\" & strFolderDate & "\AdHoc\ScheduledJobs /E /R:10 /W:30 /NP", 7, True)
  lngReturnValue = objShell.Run("robocopy.exe C:\Software\DLLs " & conCopyToFolder & "\" & strFolderDate & "\AdHoc\DLLs /E /R:10 /W:30 /NP", 7, True)
End If			

Set objSubFolders = Nothing
Set objFolder = Nothing
Set objShell = Nothing
Set objFSO = Nothing

If conDebugMode Then wscript.echo "Done"

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