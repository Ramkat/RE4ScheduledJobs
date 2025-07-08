'Filename    : DeleteATCamImages.vbs
'Author      : Christo Pretorius
'Date        : 24 June 2005
'Description : This script will retrieve the latest cam image number from the database
'            : and delete the previous 48 images (24hrs x 2/hr = 48 images)
'
'Note        : This job must run only once per day AFTER the daily website backup.

CONST conDebug = False

Call DeleteImages

If conDebug Then
	wscript.echo "Quit with success"
	wscript.Quit 0					'Quit with success
Else
	wscript.Quit 0					'Quit with success
End If

Sub DeleteImages
	Dim ADOCn
	Dim ADORs
	Dim strSQL
	Dim objFSO
	Dim intImageNumber
	Dim intEnd

	On Error Resume Next
	Err.Clear
	Set ADOCn = OpenADOConnection

	'Check if the connection opened successfully.
	If Err.number <> 0 Then		
		If conDebug Then
			wscript.echo "Failed to open a DB connection:" & vbcrlf & err.Description
		End If
		
		wscript.Quit 1					'Quit with failure
	End If
	
	'Get the current image number.
	strSQL = "sp_GetSetNextCameraImg 13" '13 = AfriTrust in table Website		
	Err.Clear 
	Set ADORs = OpenADORsReadOnly(ADOCn, strSQL, True)
	
	If Err.number <> 0 Then
		If conDebug Then
			wscript.echo "Failed to get image number:" & vbcrlf & err.Description
		End If
		
		Set ADORs = Nothing
		Set ADOCn = Nothing
		wscript.Quit 1					'Quit with failure
	End If
			
	If ADORs.RecordCount = 1 Then
		intImageNumber = CInt(ADORs("Return_Value"))
	End If

	If conDebug Then
		wscript.echo "intImageNumber = " & intImageNumber
	End If

	Set ADORs = Nothing
	Set ADOCn = Nothing
	
	Err.Clear 
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	If Err.number <> 0 Then
		Set objFSO = Nothing
		wscript.Quit 1					'Quit with failure
	End If
	
	'Delete the previous 48 images.
	intImageNumber = intImageNumber - 1
	intEnd = intImageNumber - 48
	
	Do While intImageNumber >= intEnd
		'Delete the image
		If conDebug Then
			wscript.echo "Deleting file:" & vbcrlf & "c:\www\afritrust\camera\images\" & intImageNumber & ".jpg"
		Else		
			objFSO.DeleteFile "c:\www\afritrust\camera\images\" & intImageNumber & ".jpg", True
		End If

		'Do not test for errors, just continue.
		intImageNumber = intImageNumber - 1
	Loop

	Set objFSO = Nothing
End Sub

Function OpenADOConnection
	'This function will open a connection to the database
	'and return the connection.

	OpenADOConnection = "" 	'Assume failure

	Dim ADOCn
	Dim strConnString

	Set ADOCn = CreateObject("ADODB.Connection")
	
	'ODBC Connection
	strConnString = "driver={SQL Server};pwd=~letzgetle@rn1ng!;uid=WlcUser;database=wlc;Server=127.0.0.1,11523"

	ADOCn.CommandTimeout = 60
	ADOCn.CursorLocation = 3	'Client-side cursor. DO NOT CHANGE IT !!!
	ADOCn.Open strConnString

	If Not ADOCn Is Nothing Then
		Set OpenADOConnection = ADOCn
	End If
End Function

Function OpenADORsReadOnly(ADOConnection, strSQL, blnDisconnect)
	'This funciton will return a READ ONLY, ADO recordset.
	'If blnDisconnect = True, the recordset will be disconnected.

	OpenADORsReadOnly = "" 		'Assume failure

	Dim ADORs

	Set ADORs = CreateObject("ADODB.Recordset")	
	ADORs.Open strSQL, ADOConnection, 3, 1

	If Not ADORs Is Nothing Then
		If blnDisconnect Then
			'Disconnect the recordset
			ADORs.ActiveConnection = Nothing
		End If
		
		Set OpenADORsReadOnly = ADORs		
	End If
		
	Set ADORs = Nothing
End Function