'Filename    : BackupSQL.vbs
'Author      : Christo Pretorius
'Date        : 17 September 2009
'Description : This script will copy the SQL backups to the web server.
'            : It will delete backup files and folders older than conBackupDays 
'			   and copy the latest files.

Option Explicit
On Error Goto 0

CONST conBackupFolder = "\\10.0.0.2\C$\SQLBackups"
CONST conCopyToFolder = "c:\Backups_DB1\SQLBackups"
CONST conBackupDays = 3
CONST conDebugMode = False

Dim objFSO
Dim objFolder
Dim objSubFolders
Dim objSQLFolders
Dim objShell
Dim SubFolder
Dim SQLFolder
Dim File
Dim dteDate
Dim strDate
Dim strTestDate
Dim strExtension
Dim lngReturnValue

'Create an object of the file system
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Create an object of the cmd shell
Set objShell = CreateObject("WScript.Shell")

dteDate = Date()
strDate = Trim(CStr(Year(dteDate))) & Trim(CStr(Month(dteDate))) & Trim(CStr(Day(dteDate)))

dteDate = Date() - conBackupDays
strTestDate = Trim(CStr(Year(dteDate))) & Trim(CStr(Month(dteDate))) & Trim(CStr(Day(dteDate)))

'=====
'Copy the latest backup files to the conCopyToFolder.
'=====
'Create an object of the backups folder
Set objFolder = objFSO.GetFolder(conBackupFolder)

'Create an object for all the sub folders of the backups folder
Set objSubFolders = objFolder.SubFolders

'This was the original date calculation. I changed it
'on 1 October 2004 so that only 1 day's backup is kept
'dteDate = Date - conBackupDays - 1
dteDate = Date - 2	'Use minus 2, else the latest backup gets deleted.

'Loop through all the sub folders
For Each SubFolder In objSubFolders
	
		For Each File In SubFolder.Files
			strExtension = LCase(objFSO.GetExtensionName(File.Name))

			'Check if the extension ends with .bak/.trn
			If strExtension = "bak" or strExtension = "trn" Then								 			
				If File.DateLastModified < dteDate Then
					'Delete the file if it is older than Date minus conBackupDays minus 1
					If Not conDebugMode Then
						On Error Resume Next
						objFSO.DeleteFile conBackupFolder & "\" & SubFolder.Name & "\" & File.Name
						On Error Goto 0
					Else					
						wscript.echo "Deleting file " & conBackupFolder & "\" & SubFolder.Name & "\" & File.Name
					End If
				ElseIf DateDiff("d", File.DateCreated, date) < 1 Then 
					'If the file is younger than 24 hours, copy it.

					If Not conDebugMode Then
						On Error Resume Next						
						lngReturnValue = objShell.Run("robocopy.exe " & conBackupFolder & "\" & SubFolder.Name & " " & conCopyToFolder & " " & File.Name & " /R:10 /W:30 /NP", 7, True)
'wscript.echo "robocopy.exe " & conBackupFolder & "\" & SubFolder.Name & " " & conCopyToFolder " " & File.Name & " /R:10 /W:30 /NP"
'wscript.quit
						On Error Goto 0
					Else
						wscript.echo "Copying file " & conBackupFolder & "\" & SubFolder.Name & "\" & File.Name & vbcrlf & "to " & conCopyToFolder & "\" & File.Name & vbcrlf _
							& "robocopy.exe " & conBackupFolder & "\" & SubFolder.Name & " " & conCopyToFolder & " " & File.Name & " /R:10 /W:30 /NP"
					End If							
				End If
			Else
				'The filename doesn't end with .bak/.trn -> delete it.
				If Not conDebugMode Then
					On Error Resume Next
					objFSO.DeleteFile conBackupFolder & "\" & SubFolder.Name & "\" & File.Name
					On Error Goto 0
				Else
					wscript.echo "Deleting file " & conBackupFolder & "\" & SubFolder.Name & "\" & File.Name
				End If
			End If		
		Next
Next

Set objShell = Nothing
Set objSubFolders = Nothing
Set objFolder = Nothing

'=====
'Go through conCopyTo folder and remove all files older than conBackupDays
'=====
'Create an object of the CopyTo folder
Set objFolder = objFSO.GetFolder(conCopyToFolder)

dteDate = Date - conBackupDays

'Loop through all the files
For Each File In objFolder.Files
  strExtension = LCase(objFSO.GetExtensionName(File.Name))

  'Check if the extension ends with .bak/.trn
  If strExtension = "bak" or strExtension = "trn" Then								 			
    If File.DateLastModified < dteDate Then
      'Delete the file if it is older than Date minus conBackupDays minus 1
	If Not conDebugMode Then
	  On Error Resume Next
	  objFSO.DeleteFile conCopyToFolder & "\" & File.Name
	  On Error Goto 0
	Else					
	  wscript.echo "Deleting file " & conCopyToFolder & "\" & File.Name
	End If
    End If
  End If
Next

Set objFolder = Nothing
Set objFSO = Nothing

If conDebugMode Then wscript.echo "Done"