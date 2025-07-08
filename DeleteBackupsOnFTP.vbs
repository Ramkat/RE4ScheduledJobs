'Filename    : DeleteBackupsOnFTP.vbs
'Author      : Christo Pretorius
'Date        : 18 July 2017
'Description : This script will remove backup folders from the FTP server that are older than conBackupDays.
'			 : This script is meant to be run manually.

Option Explicit

CONST conBackupFolder = "D:"
CONST conBackupDays = 7
CONST conDebugMode = False

Dim objFSO
Dim objShell

Dim objFolder
Dim objSubFolders
Dim SubFolder

Dim WebFolder
Dim dteDate
Dim strTestDate
Dim intDay
Dim fileName
Dim objFile
Dim dteFolderDate

'Create an object of the file system
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Create an object of the cmd shell
Set objShell = CreateObject("WScript.Shell")

'Create a file for logging
fileName = "DeleteBackupsOnFTP-" & strGetDateAsNumerics(Now) & ".log"
Set objFile = objFSO.OpenTextFile(fileName, 2, True)	' 2=For Writing

If conDebugMode Then Wscript.echo "Log filename = " & fileName

'Get the date less conBackupDays.
dteDate = DateAdd("d", conBackupDays * -1, Now)  'Subtract x days from today.

If conDebugMode Then Wscript.echo "dteDate = " & dteDate

strTestDate = strGetDateAsNumerics(dteDate)

If conDebugMode Then Wscript.echo "strTestDate = " & strTestDate

'=====
'Delete all the backup folders older than conBackupDays in the conBackupFolder.
'=====

'=====Temp code start
'Dim objFolderDate
'Dim objSubFoldersDate
'Dim SubFolderDate
'
'Dim objFolderWebSite
'Dim objSubFoldersWebsite
'Dim SubFolderWebsite
'=====Temp code	end
	
'Create an object of the "Backup" folder
Set objFolder = objFSO.GetFolder(conBackupFolder)

'Create an object for all the sub folders of the "Backup" folder
Set objSubFolders = objFolder.SubFolders

On Error Resume Next
	
'Loop through all the sub folders
For Each SubFolder In objSubFolders	
	dteFolderDate = CDate(strMakeDateFromString(SubFolder.Name))
	
	'This if statement ensures that we only delete valid folders.		
	If Err.Number = 0 Then	
		If conDebugMode Then Wscript.Echo "SubFolder.Name as date = " & dteFolderDate
		
		'If the web sub folder is older than conBackupDays or its name is the date minus conBackupDays, delete it.
		'If SubFolder.DateCreated < dteDate Or SubFolder.Name = strTestDate Then	
		If dteFolderDate < dteDate Or SubFolder.Name = strTestDate Then	
			If Not conDebugMode Then
				objFile.WriteLine "Deleting folder " & conBackupFolder & "\" & SubFolder.Name
				On Error Resume Next
				objFSO.DeleteFolder conBackupFolder & "\" & SubFolder.Name, True
				On Error Goto 0
			Else
				wscript.echo "Deleting " & conBackupFolder & "\" & SubFolder.Name
			End If	
		Else
			objFile.WriteLine "Skipped folder " & SubFolder.Name & ". DateCreated = " & SubFolder.DateCreated
		End If
		'=====END OF THE ORIGINAL CODE
		
		'=====TEMP CODE	START
		'Create an object of the datename folder
		'Set objFolderDate = objFSO.GetFolder(conBackupFolder & "\" & SubFolder.Name)

		'Create an object for all the sub folders of the datename folder. Thus the website names.
		'Set objSubFoldersDate = objFolderDate.SubFolders
		
		'Loop through all the sub folders of the datename folder
		'For Each SubFolderDate In objSubFoldersDate
		'	'Create an object of the website name folder
		'	If conDebugMode = True Then Wscript.Echo "GetFolder('" & conBackupFolder & "\" & SubFolder.Name & "\" & SubFolderDate.Name & "')"
		'	Set objFolderWebSite = objFSO.GetFolder(conBackupFolder & "\" & SubFolder.Name & "\" & SubFolderDate.Name)
	'
	'		'Create an object for all the sub folders of the datename folder. Thus the website names.
	'		Set objSubFoldersWebsite = objFolderWebSite.SubFolders
	'		
	'		'Loop through all the folders in the website name folder
	'		For Each SubFolderWebsite In objSubFoldersWebsite
	'			'If the folder's create date is older than conBackupDays, delete it.
	'			If SubFolderWebsite.DateCreated < dteDate Then
	'				If Not conDebugMode Then
	'			 		objFile.WriteLine "Deleting folder " & conBackupFolder & "\" & SubFolder.Name & "\" & SubFolderDate.Name & "\" & SubFolderWebsite.Name
	'			 		On Error Resume Next
	'			 		objFSO.DeleteFolder conBackupFolder & "\" & SubFolder.Name & "\" & SubFolderDate.Name & "\" & SubFolderWebsite.Name, True
	'			 		On Error Goto 0
	'		 		Else
	'					 wscript.echo "Deleting " & conBackupFolder & "\" & SubFolder.Name & "\" & SubFolderDate.Name & "\" & SubFolderWebsite.Name
	'		 		End If				
	'			End If
	'		Next
	'		
	'		Set objSubFoldersWebsite = Nothing
	'		Set objFolderWebSite = Nothing
	'	Next
	'		
	'	Set objSubFoldersDate = Nothing
	'	Set objFolderDate = Nothing
		'====TEMP CODE END
	End If
Next

On Error Goto 0
	
objFile.Close
Set objFile = Nothing
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
	strDay = Day(dteDate)			'Get the day
	
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

Function strMakeDateFromString(strDate)
	strMakeDateFromString = Mid(strDate, 1, 4) & "/" & Mid(strDate, 5, 2) & "/" & Mid(strDate, 7,2)	
End Function