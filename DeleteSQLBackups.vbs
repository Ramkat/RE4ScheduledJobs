'Filename    : DeleteSQLBackups.vbs
'Author      : Christo Pretorius
'Date        : 29 July 2010
'Description : This script will delete backup files older than conBackupDays.

Option Explicit
On Error Goto 0

CONST conBackupFolder = "D:\SQLBackups"
CONST conBackupDays = 14
CONST conDebugMode = False

Dim objFSO
Dim objFolder
Dim File
Dim strExtension
Dim dteDate

'Create an object of the file system
Set objFSO = CreateObject("Scripting.FileSystemObject")

'Create an object of the backup folder
Set objFolder = objFSO.GetFolder(conBackupFolder)

dteDate = Date - conBackupDays

'Loop through all the files
For Each File In objFolder.Files
  strExtension = LCase(objFSO.GetExtensionName(File.Name))

  'Check if the extension ends with .bak/.trn/.txt
  If strExtension = "bak" or strExtension = "trn" or strExtension = "txt" Then								 			
    If File.DateLastModified < dteDate Then
      'Delete the file if it is older than Date minus conBackupDays
		If Not conDebugMode Then
		  On Error Resume Next
		  objFSO.DeleteFile conBackupFolder & "\" & File.Name
		  On Error Goto 0
		Else					
		  wscript.echo "Deleting file " & conBackupFolder & "\" & File.Name
		End If
    End If
  End If
Next

Set objFolder = Nothing
Set objFSO = Nothing

If conDebugMode Then wscript.echo "Done"