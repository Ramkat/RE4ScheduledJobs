'Filename    : TestAzureSQL.vbs
'Created by  : Christo Pretorius	16 September 2020
'Description : This script is used to test a connection to a SQL Server on a VM hosted on Azure.
		
Dim ADOCn
Dim strSQL
Dim blnGetEmails

Call Main

Sub Main
	
	'On Error Resume Next
	Err.Clear
	
	Call TestSQL		
			
End Sub

Sub TestSQL
	'This sub will count all email
	strSQL = "select count(*) as Counted from email where emailstatusid = 1"	
	
	'Open a connection to the database.	
	'On Error Resume Next
	Err.Clear
	Set ADOCn = OpenADOConnection

	'Check if the connection opened successfully.
	If Err.number <> 0 Then
		On Error Goto 0
		Exit Sub
	End If
	
	Err.Clear 
	Set ADORs = OpenADORsReadOnly(ADOCn, strSQL, True)
	
	If Err.number <> 0 Then
		Set ADORs = Nothing
		Set ADOCn = Nothing
		Exit Sub
	End If
			
	If ADORs.RecordCount > 0 And ADORs("Counted") > 0 Then
		wscript.echo ADORs.RecordCount & " unsent records in the Email table."
	Else
		wscript.echo "No unsent records found in the Email table."
	End If		

	Set ADORs = Nothing
	Set ADOCn = Nothing
End Sub

Function OpenADOConnection
	'This function will open a connection to the database
	'and return the connection.

	OpenADOConnection = "" 	'Assume failure

	Dim ADOCn
	Dim strConnString

	Set ADOCn = CreateObject("ADODB.Connection")
	
	'20.50.111.103 is the IP address of the SQL server in Azure.
	strConnString = "Driver={SQL Server Native Client 11.0};pwd=~letzgetem@1l!;uid=EmailUser;database=wlc;Server=127.0.0.1\SQL2017"	
	ADOCn.CommandTimeout = 60
	ADOCn.CursorLocation = 3	'Client-side cursor. DO NOT CHANGE IT !!!
	ADOCn.Open strConnString

	If Not ADOCn Is Nothing Then
		Set OpenADOConnection = ADOCn
	End If
End Function

Function OpenADORsReadOnly(ADOConnection, strSQL, blnDisconnect)
	'This function will return a READ ONLY, ADO recordset.
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

Function GetSQLErrors(ADOCn)
	'This function will retrieve the native SQL errors
	'from the supplied connection.
		
	Dim intError
	Dim strReturn
	
	For intError = 0 To ADOCn.Errors.Count - 1
		'Get all the SQL errors and their descriptions.		
		strReturn = strReturn & ADOCn.Errors.Item(intError).Description & vbcrlf
	Next
		
	GetSQLErrors = strReturn
End Function