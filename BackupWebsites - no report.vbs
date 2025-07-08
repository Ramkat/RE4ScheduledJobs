'Filename    : BackupWebsites.vbs
'Author      : Christo Pretorius
'Date        : 9 September 2003
'Description : This script will copy the website folders to the backup fodler with today's date.
'            : It will delete backup files and folders older than conBackupDays 
'    	       and copy the latest files.

Option Explicit

CONST conBackupFolder = "C:\Webs"
CONST conCopyToFolder = "D:\Webs-Backups"
CONST conFoldersToSkip = "Upload WildlifeCampus-Test Updates Logs"
CONST conBackupDays = 14
CONST conDebugMode = False

Dim objFSO
Dim objFolder
Dim objSubFolders
Dim objShell
Dim SubFolder
Dim dteDate
Dim strDateToday

'Create an object of the file system
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Create an object of the "Copy To" folder
Set objFolder = objFSO.GetFolder(conCopyToFolder)

'Create an object for all the sub folders of the "Copy To" folder
Set objSubFolders = objFolder.SubFolders

'Create an object of the cmd shell
Set objShell = CreateObject("WScript.Shell")

dteDate = Date()
strDateToday = strGetDateAsNumerics(dteDate)

dteDate = Date() - conBackupDays

'=====
'Delete all the backup folders older than conBackupDays in the "Copy To" folder
'and create a new folder with today's date for the new backup files.
'=====
'Loop through all the sub folders
For Each SubFolder In objSubFolders
	'If the sub folder is older than conBackupDays or its name is older than the date minus conBackupDays, delete it.
	If SubFolder.DateCreated < dteDate Or dteGetFolderDateAsDate(SubFolder.Name) < dteDate Then
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
'Copy the website folders to the "conCopyToFolder\today's date "
'=====
													
If Not conDebugMode Then
	On Error Resume Next		
	lngReturnValue = objShell.Run("robocopy.exe " & conBackupFolder & " " & conCopyToFolder & "\" & strDateToday & " /E /R:10 /W:5 /NP /XD " & conFoldersToSkip, 7, True)
	On Error Goto 0
Else
	wscript.echo "Copying folder " & conBackupFolder & vbcrlf & "to " & conCopyToFolder & "\" & strDateToday & vbcrlf _
		& "robocopy.exe " & conBackupFolder & " " & conCopyToFolder & "\" & strDateToday & " /E /R:10 /W:5 /NP /XD " & conFoldersToSkip
End If
	

'The BackupWebsitesToFTP.vbs script doesn't want to work from here. 
'It now scheduled via Windows Scheduler.
'Start the script to copy from the website BACKUP folder to the FTP folder.
'On Error Resume Next	
'lngReturnValue = objShell.Run("wscript.exe C:\ScheduledJobs\BackupWebsitesToFTP.vbs " & strDateToday, 7, False)
'On Error Goto 0

Set objFSO = Nothing
Set objShell = Nothing

If conDebugMode Then	wscript.echo "Done"
Wscript.Quit 0

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

Function dteGetFolderDateAsDate(strFolderDate)
    ' Check if the input string is exactly 8 characters long
    If Len(strFolderDate) <> 8 Then
        dteGetFolderDateAsDate = ""
        Exit Function
    End If

    ' Extract year, month, and day from the string
    Dim year, month, day
    year = Mid(strFolderDate, 1, 4)
    month = Mid(strFolderDate, 5, 2)
    day = Mid(strFolderDate, 7, 2)

    ' Create a date string in a recognizable format
    Dim formattedDate
    formattedDate = year & "/" & month & "/" & day ' YYYY/MM/DD format

    ' Convert to date object
    On Error Resume Next ' Handle any conversion errors
    dteGetFolderDateAsDate = CDate(formattedDate)
	
    If Err.Number <> 0 Then
        dteGetFolderDateAsDate = ""
        Err.Clear
    End If
    On Error GoTo 0 ' Reset error handling
End Function