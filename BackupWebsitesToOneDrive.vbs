'Filename    : BackupWebsitesToOneDrive.vbs
'Author      : Christo Pretorius
'Date        : 3 Feb 2021
'Description : This script will copy the website backups to the OneDrive folder. 
'				Scheduled to run on 2nd and 16th of each month.

Option Explicit

'DISABLED wscript.quit 'Exit without execution

CONST conDebugMode = False '###

CONST conBackupFolder = "D:\Webs-Backups"
CONST conCopyToFolder = "D:\BackupToOneDrive"

'Note: Only RE4's websites are backed up. If client websites needs to be be backed up,
'then the "Set objFolder" cmd below needs to change since client websites are in a different folder.
CONST conWebsitesToBackup = "AdminJS,CulturalGuiding,InvestorCampus,secure.re4.co.za,WildlifeCampus"

Dim objIEDebugWindow
Dim strFolderDate
Dim strZipFolder
Dim strZipFile
Dim boolZipped
Dim objWshShell
Dim objFSO
Dim objFolder
Dim objSubFolders
Dim SubFolder
Dim WebFolder
Dim dteDate
Dim strDate
Dim strTestDate
Dim intDay
Dim dteFolderDate
Dim fileName
Dim objFile
Dim strCmd
Dim lngReturnValue

On Error Goto 0

'Create an object of the file system
Set objFSO = CreateObject("Scripting.FileSystemObject")

dteDate = Date() - 1	'### 
strFolderDate = strGetDateAsNumerics(dteDate)

On Error Resume Next 'The OpenTextFile cmd below sometimes gives a permission denied error. Ignore it.
fileName = "D:\ScheduledJobs\BackupWebsitesToOneDrive-" & strGetDateAsNumerics(dteDate) & ".log"
Set objFile = objFSO.OpenTextFile(fileName, 2, True)	' 2=For Writing
On Error Goto 0 
objFile.WriteLine "BackupWebsitesToOneDrive.vbs started at " & Now

Set objWshShell = WScript.CreateObject("WScript.Shell")

objFile.WriteLine "Parameter: " & strFolderDate
Output "Parameter: " & strFolderDate

'Create an object of the "Backups" folder
Set objFolder = objFSO.GetFolder(conBackupFolder & "\" & strFolderDate & "\RE4")

Call StartZip
Call CopyAdHocFolders

objFile.WriteLine "BackupWebsitesToOneDrive.vbs completed at " & Now
objFile.Close

Set objFile = Nothing
Set objSubFolders = Nothing
Set objFolder = Nothing
Set objFSO = Nothing
Set objWshShell = Nothing

Output "Done"
Wscript.Quit 0

Sub StartZip
	'=====
	'Zip & copy the website folders to the conCopyToFolder.
	'=====
	
	objFile.WriteLine "Starting 'StartZip'"
	Output "Starting 'StartZip'"

	'Check if the zip executable & dll exist. Must be in the same folder as *THIS* .vbs file.
	If Not objFSO.FileExists("D:\ScheduledJobs\7z.exe") Then		
		objFile.WriteLine "    7z.exe does not exist. Exiting."
		Output "    7z.exe does not exist. Exiting."
		Exit Sub
	End If
	
	If Not objFSO.FileExists("D:\ScheduledJobs\7z.dll") Then		
		objFile.WriteLine "    7z.dll does not exist. Exiting."
		Output "    7z.dll does not exist. Exiting."
		Exit Sub
	End If
			
	Dim intStopCount
	Dim lngReturnValue

	'Disabled because robocopy creates the folder.
	'Check if the target folder with the folder date passed as an argument exists.
	'If Not objFSO.FolderExists(conCopyToFolder & "\" & strFolderDate) Then
	'	If Not conDebugMode Then
	'		On Error Resume Next
	'		objFile.WriteLine "  Creating folder " & conCopyToFolder & "\" & strFolderDate
	'		objFSO.CreateFolder(conCopyToFolder & "\" & strFolderDate)
	'		On Error Goto 0
	'	Else
	'		Output "StartZip: Creating folder " & conCopyToFolder & "\" & strFolderDate
	'	End If
	'End If

	'Create an object for all the sub folders of the backups folder (the website names).
	Set objSubFolders = objFolder.SubFolders
	
	intStopCount = 0	'Use with conDebugMode to test just a few folders.

	'Loop through all the sub folders of the backups folder (the website names)
	For Each SubFolder In objSubFolders
		intStopCount = intStopCount + 1
		
		objFile.WriteLine "Processing folder: " & SubFolder.Name
		Output "Processing folder: " & SubFolder.Name
		
		If intStopCount = 5 And conDebugMode = True Then
			objFile.WriteLine "5 file test complete. Exiting StartZip."
			Output "5 file test complete. Exiting StartZip."
			Exit Sub
		End If
		
		'Check if this folder needs to be processed.
		If InStr(conWebsitesToBackup, SubFolder.Name) > 0 Then
			objFile.WriteLine "Zip & move folder: " & SubFolder.Name
			Output "Zip & move folder: " & SubFolder.Name
		
			strZipFolder = conBackupFolder & "\" & strFolderDate & "\RE4\" & SubFolder.Name
			strZipFile =  strZipFolder & "-" & strFolderDate & ".zip"	

			'Check if the zipped folder's file exists. If not, then zip it.
			'Reason: if a backup failed, this script can be run again without creating duplicate zip files. Saves a LOT of time.
			If objFSO.FileExists(strZipFile) = False Then
				boolZipped = DoZip(strZipFolder, strZipFile)
			Else
				boolZipped = True
			End If
			
			strCmd = "robocopy.exe " & conBackupFolder & "\" & strFolderDate & "\RE4" & " " & conCopyToFolder & "\" & strFolderDate & " " & SubFolder.Name & "-" & strFolderDate & ".zip /MOV /R:10 /W:30 /NP"
																	
			If Not conDebugMode Then
				On Error Resume Next
						
				'If the zip command was successful, move the file.
				If boolZipped = True Then						
					objFile.WriteLine "    Moving file " & strZipFile & " to " & conCopyToFolder & "\" & strFolderDate
					objFile.WriteLine "    Cmd: " & strCmd
					
					lngReturnValue = objWshShell.Run(strCmd, 7, True)	'Move the file with RoboCopy				
					objFile.WriteLine "    Zip file move return value = " & lngReturnValue
					
					'Delete the zipped backup file if it still exists.
					If objFSO.FileExists(strZipFile) Then
						objFSO.DeleteFile(strZipFile)
					End If
				End If
				
				On Error Goto 0
			Else						
				Output "StartZip: Moving file " & strZipFile & vbcrlf & "to " & conCopyToFolder & "\" & strFolderDate & vbcrlf & strCmd
			End If
			
			objFile.WriteLine "    - - - - -"
		End If		
	Next
	
	objFile.WriteLine "Exiting 'StartZip'" & vbCrLf
End Sub

Function DoZip(strFolderPathAndNameToZip, strZipFilePathAndName)			    
    If Not objFSO.FolderExists(strFolderPathAndNameToZip) Then		
		Output "DoZip: Folder " & strFolderPathAndNameToZip & " to zip does not exist."
		objFile.WriteLine "    Folder " & strFolderPathAndNameToZip & " to zip does not exist."
			
		DoZip = False
		Exit Function
    End If
	
	If objFSO.FileExists(strZipFilePathAndName) Then
		Output "DoZip: File " & strZipFilePathAndName & " already exists - deleting it."
		objFile.WriteLine "DoZip: File " & strZipFilePathAndName & " already exists - deleting it."
		objFSO.DeleteFile strZipFilePathAndName		
    End If
        
    Dim lngReturnValue
	Dim strCmd
    strCmd = "D:\ScheduledJobs\7z.exe a " & strZipFilePathAndName & " " & strFolderPathAndNameToZip & " -mx=5 -ssw"	
	'Command switches:  a -> add ; -mx=5 -> compression level 5 of 9 ; -ssw -> compress files open for writing
	
	If conDebugMode Then
		Output "DoZip: Zip cmd = " & strCmd
	Else
		objFile.WriteLine "    Zipping: " & strCmd
		'Execute the command and wait for it to finish.
		lngReturnValue = objWshShell.Run(strCmd, 7, True) '7=Displays the window as a minimized window. The active window remains active.		
		objFile.WriteLine "    Zipping: Return value = " & lngReturnValue
		
		If lngReturnValue = 0 Then
			DoZip = True
		Else
			DoZip = False
		End If
	End If
End Function

Sub CopyAdHocFolders
	'Backup a few ad-hoc folders	
	
	On Error Resume Next
	objFile.WriteLine "Copy ad-hoc folders"
	
	'=====
	'MiscInfo Folder
	'=====							
	strZipFolder = "D:\MiscInfo"
	strZipFile = "D:\Temp\MiscInfo-" & strFolderDate & ".zip"	

	'Check if the zipped folder's file exists. If not, then zip it.
	'Reason: if a backup failed, this script can be run again without creating duplicate zip files. Saves a LOT of time.
	If objFSO.FileExists(strZipFile) = False Then
		boolZipped = DoZip(strZipFolder, strZipFile)
	Else
		boolZipped = True
	End If
				
	strCmd = "robocopy.exe D:\Temp " & conCopyToFolder & "\" & strFolderDate & " MiscInfo-" & strFolderDate & ".zip /MOV /R:10 /W:30 /NP"
	
	If Not conDebugMode Then
		objFile.WriteLine "    Cmd: " & strCmd
		lngReturnValue = objWshShell.Run(strCmd, 7, True)		
		objFile.WriteLine "    D:\MiscInfo folder zip & move return value = " & lngReturnValue  
	End If
		
	Output "    Cmd: " & strCmd
	
	'=====
	'ScheduledJobs Folder
	'=====	
	strZipFolder = "D:\ScheduledJobs"
	strZipFile = "D:\Temp\ScheduledJobs-" & strFolderDate & ".zip"	

	'Check if the zipped folder's file exists. If not, then zip it.
	'Reason: if a backup failed, this script can be run again without creating duplicate zip files. Saves a LOT of time.	
	
	If objFSO.FileExists(strZipFile) = False Then
		boolZipped = DoZip(strZipFolder, strZipFile)
	Else
		boolZipped = True
	End If
		
	strCmd = "robocopy.exe D:\Temp " & conCopyToFolder & "\" & strFolderDate & " ScheduledJobs-" & strFolderDate & ".zip /MOV /R:10 /W:30 /NP"
		
	If Not conDebugMode Then	
		objFile.WriteLine "    Cmd: " & strCmd
		lngReturnValue = objWshShell.Run(strCmd, 7, True)
		objFile.WriteLine "    D:\ScheduledJobs folder zip & move return value = " & lngReturnValue  
	End If

	Output "    Cmd: " & strCmd
	
	'=====
	'SQLBackups Folder
	'=====				
	'Create a wildcard filename to use, based on the date.
	Dim strSQLDate
	strSQLDate = Left(strFolderDate, 4) & "_" & Mid(strFolderDate, 5,2) & "_" & Mid(strFolderDate, 7)
	
	strCmd = "robocopy.exe D:\SQLBackups " & conCopyToFolder & "\" & strFolderDate & "\SQLBackups *" & strSQLDate & "*.bak /R:10 /W:30 /NP"
	
	If Not conDebugMode Then
		objFile.WriteLine "    Cmd: " & strCmd
		lngReturnValue = objWshShell.Run(strCmd, 7, True)		
		objFile.WriteLine "    D:\SQLBackups folder copy return value = " & lngReturnValue  
	End If
		
	Output "    Cmd: " & strCmd
	
	'=====
	EXIT SUB
	'DLL files no longer in use.
	'=====
	
	strZipFolder = "D:\Software\DLLs"
	strZipFile = "D:\Temp\DLLs-" & strFolderDate & ".zip"	

	'Check if the zipped folder's file exists. If not, then zip it.
	'Reason: if a backup failed, this script can be run again without creating duplicate zip files. Saves a LOT of time.
	If objFSO.FileExists(strZipFile) = False Then
		boolZipped = DoZip(strZipFolder, strZipFile)
	Else
		boolZipped = True
	End If

	lngReturnValue = objWshShell.Run("robocopy.exe " & strZipFile & " " & conCopyToFolder & "\" & strFolderDate & "\AdHoc /MOV /R:10 /W:30 /NP", 7, True)
	objFile.WriteLine "    D:\Doftware\DLLs folder zip & move return value = " & lngReturnValue  
	
	On Error Goto 0

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

Function strMakeDateFromString(strDate)
	strMakeDateFromString = Mid(strDate, 1, 4) & "/" & Mid(strDate, 5, 2) & "/" & Mid(strDate, 7,2)	
End Function

Sub Output(myText)
	If conDebugMode = False Then Exit Sub	'Ensure that if we don't show unnecessary windows if debugging isn't enabled.

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