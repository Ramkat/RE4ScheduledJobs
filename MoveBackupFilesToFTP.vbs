'Filename    : MoveBackupFilesToFTP.vbs
'Author      : Christo Pretorius
'Date        : 3 Aug 2017
'Description : This script will loop through the backups folder and move all zip files to the backup drive to the FTP server.

Option Explicit

'DISABLED wscript.quit 'Exit without execution

CONST conBackupFolder = "C:\Webs-Backup"
CONST conCopyToFolder = "D:"
CONST conDebugMode = False

Dim strFolderDate
Dim objWshShell
Dim objFSO
Dim objFolder
Dim objSubFolders
Dim SubFolder
Dim strFolderToCreate
Dim File
Dim arrFile
Dim WebFolder
Dim strDate
Dim intDay
Dim dteFolderDate
Dim fileName
Dim objFile
Dim strSourceFolder
Dim strTargetFolder
Dim strCmd
Dim lngReturnValue

'Create an object of the file system
Set objFSO = CreateObject("Scripting.FileSystemObject")

fileName = "MoveBackupFilesToFTP-" & strGetDateAsNumerics(Now) & ".log"
Set objFile = objFSO.OpenTextFile(fileName, 2, True)	' 2=For Writing

objFile.WriteLine "MoveBackupFilesToFTP.vbs started at " & Now

Set objWshShell = WScript.CreateObject("WScript.Shell")

Call StartMove

Set objSubFolders = Nothing
Set objFolder = Nothing
Set objFSO = Nothing
Set objWshShell = Nothing

objFile.WriteLine "MoveBackupFilesToFTP.vbs completed at " & Now
objFile.Close
Set objFile = Nothing

If conDebugMode Then wscript.echo "Done"
Wscript.Quit

Sub StartMove
	'Create an object of the backups folder.
	Set objFolder = objFSO.GetFolder(conBackupFolder)

	'Create an object for all the sub folders of the backups folder (the website names).
	Set objSubFolders = objFolder.SubFolders

	Dim intStopCount
	intStopCount = 0	'Use with conDebugMode to test just a few folders.

	'Loop through all the sub folders of the backups folder (the website names)
	For Each SubFolder In objSubFolders
		intStopCount = intStopCount + 1
		
		If intStopCount = 3 And conDebugMode = True Then
			objFile.WriteLine "3 file test complete. Exiting StartMove."
			Exit Sub
		End If
		
		'Loop through each file in the root of the web site name folder
		For Each File in SubFolder.Files	
			'Split the filename on the dot
			arrFile = Split(File.Name, ".")
		
			'If the filename ends with zip.
			If arrFile(1) = "zip" Then							
				'Check if the target folder with the folder date (file name) exists.
				strFolderToCreate = conCopyToFolder & "\" & arrFile(0)
				If Not objFSO.FolderExists(strFolderToCreate) Then
					If Not conDebugMode Then
						On Error Resume Next
						objFile.WriteLine "  Creating folder " & strFolderToCreate
						objFSO.CreateFolder(strFolderToCreate)
						On Error Goto 0
					Else
						wscript.echo "MoveZip: Creating folder " & strFolderToCreate
					End If
				End If
				
				'Check if the target folder with the website name exists.
				strFolderToCreate = strFolderToCreate & "\" & SubFolder.Name
				If Not objFSO.FolderExists(strFolderToCreate) Then
					If Not conDebugMode Then
						On Error Resume Next
						objFile.WriteLine "  Creating folder " & strFolderToCreate
						objFSO.CreateFolder(strFolderToCreate)
						On Error Goto 0
					Else
						wscript.echo "MoveZip: Creating folder " & strFolderToCreate
					End If
				End If								
								
				strSourceFolder = conBackupFolder & "\" & SubFolder.Name 'E.g. c:\webs-backup\WildlifeCampus
				strTargetFolder = strFolderToCreate 'E.g. d:\20170802\WildlifeCampus
				
				strCmd = "robocopy.exe " & strSourceFolder & " " & strTargetFolder & " " & File.Name & " /MOV /R:10 /W:30 /NP"
																		
				If Not conDebugMode Then
					On Error Resume Next
					
					objFile.WriteLine "    Moving file " & strSourceFolder & "\" & File.Name & " to " & strTargetFolder
					'objFSO.MoveFile strSourceFolder & File.Name, strTargetFolder & "\" & File.Name
									
					lngReturnValue = objWshShell.Run(strCmd, 7, True)				
					objFile.WriteLine "    Zip file move return value = " & lngReturnValue					
					
					On Error Goto 0
				Else		
					'wscript.echo "Copying folder " & conBackupFolder & "\" & SubFolder.Name & "\" & strFolderDate & vbcrlf & "to " & conCopyToFolder & "\" & strFolderDate & "\" & SubFolder.Name & vbcrlf _
					'	& "robocopy.exe " & conBackupFolder & "\" & SubFolder.Name & "\" & strFolderDate & " " & conCopyToFolder & "\" & strFolderDate & "\" & SubFolder.Name & " /E /R:10 /W:30 /NP"
						
					wscript.echo "MoveZip: Moving file " & strSourceFolder & File.Name & " to " & strTargetFolder & vbcrlf & strCmd
				End If
				
				objFile.WriteLine "    - - - - -"
			End If 'arrFile[2] == "zip"
		Next 'For Each File in SubFolder.Files		
	Next 'For Each SubFolder In objSubFolders
	
	objFile.WriteLine "Exiting 'MoveZip'" & vbCrLf
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