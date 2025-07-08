'Filename    : SendEmailFromTable.vbs
'Created by  : Christo Pretorius	2 September 2003
'Description : This script is used to send email from the Email table
'            : on the SQL server.

Dim objSendMail
Dim blnGetEmails
Dim intEmails
Dim intLoop
Dim arrData
Dim strResult1		'Success
Dim strResult4		'Error in email address
Dim strResult9		'Unknown error
Dim ADOCn
Dim strSQL

Call Main

Sub Main
	Dim strFrom	
	Dim strTo
	
	Call GetEmails
		
	Do While blnGetEmails		
		'Loop through the email array		
		For intLoop = 0 To intEmails			
			'If the recipient address isn't empty...
		  If Len(Trim(arrData(5, intLoop))) > 0 Then 
				On Error Resume Next
				Err.Clear
				'Initialise the SMTP mailer object
				Set objSendMail = CreateObject("CDONTS.NewMail")

				If Err.Number <> 0 Then
					Set objSendMail = Nothing					
					Exit Sub
				End If
				
				On Error Goto 0
				Err.Clear
				
				'Set the FromName and FromAddress
 				strFrom = chr(34) & arrData(2, intLoop) & chr(34) & " " & arrData(3, intLoop) 				
 				objSendMail.From = strFrom
				
				'Set the recipients' address
				strTo = chr(34) & arrData(4, intLoop) & chr(34) & " " & arrData(5, intLoop)				
				objSendMail.To = strTo 				
				
				objSendMail.Subject = arrData(1, intLoop)
        objSendMail.Body = arrData(6, intLoop)
        
        'Determine the importance of the message.
        Select Case arrData(7, intLoop)
					Case 1-3	'High
						objSendMail.Importance = 2
						
					Case 4-6	'Normal
						objSendMail.Importance = 1
						
					Case 7-10	'Low
						objSendMail.Importance = 0
					
					Case Else	'Default is normal
						objSendMail.Importance = 1
				End Select
				
				'Test if it is plain or html email
				If arrData(8, intLoop) = 1 Then	
					'Text				
					objSendMail.MailFormat = 1
					objSendMail.BodyFormat = 1
				Else
					'HTML
					objSendMail.MailFormat = 0
					objSendMail.BodyFormat = 0
				End If
				
				On Error Resume Next
				Err.Clear 
				objSendMail.Send
							
				If Err.number <> 0 Then
					strResult9 = strResult9 & "," & arrData(0, intLoop)
				Else
					strResult1 = strResult1 & "," & arrData(0, intLoop)
				End If
				
				On Error Goto 0
				Set objSendMail = Nothing
		  Else
				strResult4 = strResult4 & "," & arrData(0, intLoop)
		  End If
		Next

		If UpdateResults Then
			Call GetEmails
		Else
			blnGetEmails = False
		End If
	Loop		
End Sub

Sub GetEmails
	'This sub will retrieve unsent emails.
	blnGetEmails = False
	intEmails = 0
	strSQL = "sp_GetEmailToSend"	
	
	'Open a connection to the database.	
	On Error Resume Next
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
			
	If ADORs.RecordCount > 0 Then
		arrData = ADORs.GetRows					'Put data in a 2 dimensional array
		intEmails = UBound(arrData, 2)	'Get the number of emails
		blnGetEmails = True
	End If

	Set ADORs = Nothing
	Set ADOCn = Nothing
End Sub

Function UpdateResults
	'This sub will update the result of the email send.
	
	UpdateResults = False		'Assume failure
	
	'Open a connection to the database.	
	On Error Resume Next	
	Err.Clear
	Set ADOCn = OpenADOConnection

	'Check if the connection opened successfully.
	If Err.number <> 0 Then
		On Error Goto 0
		Exit Function
	End If		
		
	If strResult1 <> "" Then
		'Remove the 1st comma
		strResult1 = Mid(strResult1, 2)
		strSQL = "exec sp_UpdateEmailResult '" & strResult1 & "', 1 ; "
	End If
		
	If strResult4 <> "" Then
		'Remove the 1st comma
		strResult4 = Mid(strResult4, 2)
		strSQL = strSQL & "exec sp_UpdateEmailResult '" & strResult4 & "', 4 ; "
	End If
	
	If strResult9 <> "" Then
		'Remove the 1st comma
		strResult9 = Mid(strResult9, 2)
		strSQL = strSQL & "exec sp_UpdateEmailResult '" & strResult9 & "', 9 ;"
	End If		
	
	On Error Resume Next	
	Err.Clear 	
	Set ADORs = OpenADORsReadOnly(ADOCn, strSQL, True)
	
	If Err.number <> 0 Then
		Set ADORs = Nothing
		Set ADOCn = Nothing
		On Error Goto 0
		Exit Function
	End If
	
	On Error Goto 0
	Set ADORs = Nothing	'Disconnect recordset
	strResult1 = ""
	strResult4 = ""
	strResult9 = ""
	UpdateResults = True
End Function

Function OpenADOConnection
	'This function will open a connection to the database
	'and return the connection.

	OpenADOConnection = "" 	'Assume failure

	Dim ADOCn
	Dim strConnString

	Set ADOCn = CreateObject("ADODB.Connection")
	
	'ODBC Connection
	strConnString = "driver={SQL Server};pwd=~letzgetem@1l!;uid=EmailUser;database=wlc;Server=10.0.0.2"

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