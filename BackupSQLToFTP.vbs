'Filename    : BackupSQLToFTP.vbs
'Author      : Christo Pretorius
'Date        : 28 October 2010
'            : 15 Mar 2018 - Added FTPUse.
'Description : This script will copy the SQL backups to the FTP server.
'	     : Because the SQL backups run at 2am, the backups will be copied
'	     : to a folder with YESTERDAY's date, because the DB backup is in 
'	     : fact the previous day's data.

'ftpuse d: ftpbackup6.jnb1.host-h.net H3tz_FtP /user:pri004_RE4_2

Option Explicit

CONST conBackupFolder = "C:\SQLBackups"
CONST conCopyToFolder = "D:"
CONST conPurgeFolder = "D:\Purge"
CONST conBackupDays = 6
CONST conDebugMode = False

Dim strFolderDate
Dim dteTestDate
Dim objAppShell 
Dim objFSO
Dim objFolder
Dim objSubFolders
Dim objShell
Dim File
Dim strExtension
Dim lngReturnValue

'Get today's date - 1 day in the format yyyymmdd
strFolderDate = strGetDateAsNumerics(Date() - 1)
dteTestDate = Date()

If conDebugMode Then Output "strFolderDate = " & strFolderDate

Set objAppShell = wscript.createobject("Shell.Application")

'Create an object of the file system
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Mount the backup drive
If Not conDebugMode Then
  If objFSO.DriveExists("d:") = False Then
    On Error Resume Next        
    objAppShell.ShellExecute "ftpuse.exe", "d: ftpbackup6.jnb1.host-h.net H3tz_FtP /user:pri004_RE4_2", "", "", 7
    wscript.Sleep 5000 'Give it 5 seconds to mount the drive.
    On Error Goto 0  
  End If
End If

Call DeleteOldBackups

'Create an object of the "Copy To" folder
Set objFolder = objFSO.GetFolder(conCopyToFolder)

'Create an object of the cmd shell
Set objShell = CreateObject("WScript.Shell")

'Check if the target folder exists.
If Not objFSO.FolderExists(conCopyToFolder & "\" & strFolderDate & "\SQLBackups") Then
	If Not conDebugMode Then
		On Error Resume Next
		objFSO.CreateFolder(conCopyToFolder & "\" & strFolderDate & "\SQLBackups")
		On Error Goto 0
	Else
		Output "Creating folder " & conCopyToFolder & "\" & strFolderDate & "\SQLBackups"
	End If
End If

'Create an object of the backups folder
Set objFolder = objFSO.GetFolder(conBackupFolder)

'Loop through all the files in the backups folder
For Each File In objFolder.Files
	strExtension = LCase(objFSO.GetExtensionName(File.Name))

	'Check if the extension ends with .bak/.trn
	If strExtension = "bak" or strExtension = "trn" Then
		'If the file is younger than yesterday's date.
		If DateDiff("d", File.DateCreated, dteTestDate) < 1 Then										
  			If Not conDebugMode Then
				On Error Resume Next
				lngReturnValue = objShell.Run("robocopy.exe " & conBackupFolder & " " & conCopyToFolder & "\" & strFolderDate & "\SQLBackups " & File.Name & " /E /R:10 /W:30 /NP", 7, True)
				On Error Goto 0
			Else
				Output "Copying file " & conBackupFolder & "\" & File.Name& " to " & vbcrlf & conCopyToFolder & "\" & strFolderDate & "\SQLBackups\" & File.Name _
				& vbcrlf & "robocopy.exe " & conBackupFolder & " " & conCopyToFolder & "\" & strFolderDate & "\SQLBackups " & File.Name & " /E /R:10 /W:30 /NP"
			End If
		End If
	End If
Next

'Dismount the backup drive
If Not conDebugMode Then
  If objFSO.DriveExists("d:") = True Then
    On Error Resume Next        
    objAppShell.ShellExecute "ftpuse.exe", "d: /delete", "", "", 7
    On Error Goto 0
  
  End If
End If			

Set objFolder = Nothing
Set objShell = Nothing
Set objFSO = Nothing
Set objAppShell = Nothing

If conDebugMode Then Output "Done"

Wscript.Quit

Sub DeleteOldBackups
	Dim dtDate
	Dim dteFolderDate 
	Dim SubFolder

	Set objFolder = objFSO.GetFolder(conBackupFolder)	

	'The purge folder must exist, else robocopy doesn't delete the files & folders.
	If objFSO.FolderExists(conPurgeFolder) = False Then 
		objFSO.CreateFolder conPurgeFolder
	End If
	
	On Error Resume Next

	'=====
	'Delete all the backup folders older than conBackupDays in the "Copy To" folder.
	'=====
	'Loop through all the sub folders
	For Each SubFolder In objSubFolders
		dteFolderDate = CDate(strMakeDateFromString(SubFolder.Name))
		
		'This if statement ensures that we only delete valid folders.		
		If Err.Number = 0 Then		
			If conDebugMode Then Output "DeleteOldBackups: SubFolder.Name as date = " & dteFolderDate
			
			'If the web sub folder is older than conBackupDays or its name is older than the date minus conBackupDays, delete it.		
			If dteFolderDate < dteDate Or SubFolder.Name = strTestDate Then
				If Not conDebugMode Then					
					'We use robocopy to delete the files and sub folders - it works the best.				
					lngReturnValue = objWshShell.Run("robocopy " & conPurgeFolder & " " & conCopyToFolder & "\" & SubFolder.Name & " /PURGE", 7, True)
					
					If lngReturnValue = 0 Then
						'Now that the folder is empty, delete the actual folder
						objFSO.DeleteFolder conCopyToFolder & "\" & SubFolder.Name, True
					End If
									
				Else
					Output "DeleteOldBackups: Deleting " & conCopyToFolder & "\" & SubFolder.Name
				End If		
			End If
		End If
	Next	
		
	On Error Goto 0
	Set objSubFolders = Nothing	
	Set objFolder = Nothing
End Sub

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

Sub Output(myText)
	If blnDebug = False Then Exit Sub	'Ensure that if we don't show unnecessary windows if debugging isn't enabled.

	If Not IsObject( objIEDebugWindow ) Then
		On Error Resume Next
		Err.Clear
		Set objIEDebugWindow = CreateObject( "InternetExplorer.Application" )
		
		If Err.Number <> 0 Then			
			wscript.echo myText
			On Error Goto 0
			Exit Sub
		End If
		
		objIEDebugWindow.Navigate "about:blank"
		objIEDebugWindow.Visible = True
		objIEDebugWindow.ToolBar = False
		objIEDebugWindow.Width   = 500
		objIEDebugWindow.Height  = 300
		objIEDebugWindow.Left    = 10
		objIEDebugWindow.Top     = 10
		
		Do While objIEDebugWindow.Busy
			WScript.Sleep 100
		Loop
		
		objIEDebugWindow.Document.Title = WScript.ScriptFullname & " output window"
		objIEDebugWindow.Document.Body.InnerHTML = "<u>" & WScript.ScriptFullname & " Output Window</u></br><br>"
	End If

	objIEDebugWindow.Document.Body.InnerHTML = objIEDebugWindow.Document.Body.InnerHTML & "<b>" & Now & ":</b>&nbsp;&nbsp;&nbsp;" & Replace(myText, vbCrLf, "<br>") & "<br>" & vbCrLf
End Sub