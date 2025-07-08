'Filename    : SendEmailFromTable.vbs
'Created by  : Christo Pretorius	2 September 2003
'Description : This script is used to send email from the Email table
'            : on the SQL server.
'			 : Added MX lookup on 14 July 2005.
'			 : Changed to CDO object instead of CDONTS on 19 July 2005.
'			 : Added file logging of sql error when updating the statuses fails on 19 July 2005.
		
Dim objSendMail
Dim blnGetEmails
Dim intEmails
Dim intLoop
Dim arrData
Dim strResult1		'Success
Dim strResult4		'Error in email address
Dim strResult7		'Host name not recognized.
Dim strResult9		'Unknown error
Dim ADOCn
Dim strSQL
Dim objIEDebugWindow
Dim blnDebug

blnDebug = false '###

Call Main

Sub Main
	Dim strFrom	
	Dim strTo
	
	On Error Resume Next
	Err.Clear
	
	'Initialise the SMTP mailer object
	Set objSendMail = CreateObject("CDO.Message")
		
	If Err.Number <> 0 Then
		If blnDebug Then Output "No email object created." & vbcrlf & Err.Description
		
		Set objSendMail = Nothing					
		Exit Sub
	End If		
				
	On Error Goto 0
	Err.Clear
	
	'Set a few values regarding the SMTP server, port etc.
'	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1 'Send the message using the local SMTP server.
'	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False 'Use SSL for the connection (True or False)
'	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

'Sending using hMailServer via re4.your-server.co.za
'	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'Send the message using a network SMTP server.
'	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "127.0.0.1"
'	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 225 '587
'	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False 'Use SSL for the connection (True or False)
'	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
'	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "webmaster@re4.your-server.co.za"
'	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "TEFqzUZb"

'Sending using Global Micro's server
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'Send the message using a network SMTP server.
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "cluster1out.eu.messagelabs.com" '"154.66.66.123"
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 '465(SSL) 
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False 'Use SSL for the connection (True or False)
	''objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendtls") = True 'Use SSL for the connection (True or False)
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "info@wildlifecampus.com"
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "W0J!s74X"		
	
'Sending using our IIS SMTP server
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'Send the message using a network SMTP server.
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "196.22.138.230" '"154.66.66.123"
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 225 '465(SSL) 
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False 'Use SSL for the connection (True or False)	
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	
	objSendMail.Configuration.Fields.Update	
	
	Call GetEmails		
		
	Do While blnGetEmails						
		'Loop through the email array		
		For intLoop = 0 To intEmails			
			'If the recipient address isn't empty...
		  If Len(Trim(arrData(5, intLoop))) > 0 Then 	
		  
				'If it is not a dummy email...
				If LCase(arrData(3, intLoop)) <> "dummy@dummy.dummy" Then
				
					'If the email's host exists, then send it
					If MXLookup(arrData(5, intLoop)) Then
						'Set the FromName and FromAddress
 						strFrom = chr(34) & arrData(2, intLoop) & chr(34) & " " & arrData(3, intLoop)
 						objSendMail.From = strFrom
				
						'Set the recipients' address
						strTo = chr(34) & arrData(4, intLoop) & chr(34) & " " & arrData(5, intLoop)				
						objSendMail.To = strTo 				
				
						objSendMail.Subject = arrData(1, intLoop)
												        						
						'Test if it is plain or html email
						If arrData(8, intLoop) = 1 Then	
							'Text				
							objSendMail.TextBody = arrData(6, intLoop)
						Else
							'HTML
							objSendMail.HtmlBody = arrData(6, intLoop)
						End If
				
						'NEWSRV  On Error Resume Next
						Err.Clear 
						objSendMail.Send
									
						If Err.number <> 0 Then
							strResult9 = strResult9 & "," & arrData(0, intLoop)
							Call LogSendError2File(arrData(0, intLoop), Err.Description)
						Else
							strResult1 = strResult1 & "," & arrData(0, intLoop)
						End If
					
						On Error Goto 0
					Else
						'The email host could not be found in a DNS lookup.
						'Thus flag it as host not recognized.
						strResult7 = strResult7 & "," & arrData(0, intLoop)
						
						If blnDebug Then 
							Output "Mailhost not found - " & arrData(5, intLoop)
							Output strResult7
						End If
					End If
				Else
					'Since it was a dummy email, flag it as successful.
					strResult1 = strResult1 & "," & arrData(0, intLoop)
				End If																
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
	
	Set objSendMail = Nothing	
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
	
	If blnDebug Then Output intEmails & " to process"

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
	strSQL = ""

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
	
	If strResult7 <> "" Then
		'Remove the 1st comma
		strResult7 = Mid(strResult7, 2)
		strSQL = strSQL & "exec sp_UpdateEmailResult '" & strResult7 & "', 7 ; "
	End If
	
	If strResult9 <> "" Then
		'Remove the 1st comma
		strResult9 = Mid(strResult9, 2)
		strSQL = strSQL & "exec sp_UpdateEmailResult '" & strResult9 & "', 9 ;"
	End If		
	
	If blnDebug Then Output strSQL
	
	On Error Resume Next	
	Err.Clear 	
	Set ADORs = OpenADORsReadOnly(ADOCn, strSQL, True)
	
	If Err.number <> 0 Then
		Call LogError2File(Err.Description)
		Set ADORs = Nothing
		Set ADOCn = Nothing
		On Error Goto 0
		Exit Function
	End If
	
	On Error Goto 0
	Set ADORs = Nothing	'Disconnect recordset
	strResult1 = ""
	strResult4 = ""
	strResult7 = ""
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
	'### dev server strConnString = "driver={SQL Server};pwd=letzgetem@1l;uid=EmailUser;database=wlc;Server=10.5.200.16"
	'strConnString = "driver={SQL Server};pwd=~letzgetem@1l!;uid=EmailUser;database=wlc;Server=127.0.0.1\SQL2017"
	
	strConnString = "Driver={SQL Server Native Client 11.0};pwd=~letzgetem@1l!;uid=EmailUser;database=wlc;Server=127.0.0.1\SQL2017"
	'Azure strConnString = "Driver={SQL Server Native Client 11.0};pwd=~letzgetem@1l!;uid=EmailUser;database=wlc;Server=20.50.111.103"		

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

Function GetEmailHost(strEmail)	
	Dim intPos
	Dim strHost
			
	intPos = 0
	intPos = InStr(1, strEmail, "@")

	strHost = ""
	strHost = Mid(strEmail, intPos + 1)
	
	'Shortest possible host is 1.tv thus if the host isn't at least 4 chars long, quit.
	If Len(strHost) < 4 Then
		GetEmailHost = ""		
	Else
		GetEmailHost = strHost
	End If
End Function

Function MXLookup(strEmail)

'Disabled on 1 Nov 2007 when the new palaeontology course
'launch email was send.
MXLookup = True
Exit Function

	'This function will try to validate the email host.
	'Upon success, it returns an empty string, else an error message.
	
  MXLookUp = False	'Assume host is not found.
  
  Dim objXMLHTTP
  Dim strResult
  Dim strHost
  
  'Get the email address's host.
  strHost = GetEmailHost(strEmail)  
  If strHost = "" Then Exit Function
  
  On Error Resume Next
  Err.Clear 
  Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
  
  If Err.number <> 0 Then  
		MXLookup = True	'Other error, allow to continue.
		Set objXMLHTTP = Nothing
		
		If blnDebug Then Output "Cannot create XMLHTTP object" & vbcrlf & Err.Description
		Exit Function
	End If
	
	If blnDebug Then Output "Email host = " & strHost
  
  objXMLHTTP.Open "Get", "http://examples.softwaremodules.com/IntraDns.asp?domainname=" & strHost & "&Submit=Submit&t_mx=1", False
  objXMLHTTP.Send
  strResult = objXMLHTTP.ResponseText    
  
  If Len(strResult) = 0 Then
		MXLookup = True	'Unable to open lookup URL, allow to continue.
		Set objXMLHTTP = Nothing				
		Exit Function
	End If  
    
  'Check the result.
	If InStr(1, strResult, "Server.CreateObject Failed") > 0 Then
		'There's a fault on the page, thus continue as if the host was found.
		MXLookup = True
	ElseIf Instr(1, strResult, "Problem Connecting to DNS Server") > 0 Then
		'DNS error, thus continue as if the host was found.
		MXLookup = True
	ElseIf Instr(1, strResult, "Unknown DNS Host") > 0 Then
		'DNS error, thus continue as if the host was found.
		MXLookup = True
	ElseIf Instr(1, strResult, "Connection Timed Out") > 0 Then
		'DNS error, thus continue as if the host was found.
		MXLookup = True
	ElseIf Instr(1, strResult, "No answer from DNS Server") > 0 Then
		'DNS error, thus continue as if the host was found.
		MXLookup = True
	ElseIf Instr(1, strResult, "DNS Server not specified") > 0 Then
		'DNS error, thus continue as if the host was found.
		MXLookup = True
	ElseIf Instr(1, strResult, "Query Timed Out") > 0 Then
		'DNS error, thus continue as if the host was found.
		MXLookup = True
	Else			
		strResult = Mid(strResult, InStr(1, strResult,"</strong>. Items Returned: <strong>") + 35, 1)
  
		If blnDebug Then Output "MxLookup result = " & strResult
  
		If CInt(strResult) > 0 Then
		  MXLookup = True	'Host found!
		  
		  If blnDebug Then Output "Mailhost found"
		End If
	End If 
  
  Set objXMLHTTP = Nothing
End Function

Sub LogSendError2File(strEmailID, strErrDesc)
	'This function will log the SMTP errors to file.
	
	'If an error occurs at this stage, ignore it.
	On Error Resume Next		
	
	Dim objFSO	
	Dim strFileName
		
	Set objFSO = CreateObject("Scripting.FileSystemObject")	
	Set objFile = objFSO.OpenTextFile("EmailSMTPErrors.txt", 8, True)
	
	objFile.WriteLine "Date : " & Now() 
	objFile.WriteLine "EmailID : " & strEmailID
	objFile.WriteLine "SMTP Error : " & strErrDesc	
	objFile.WriteLine "= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = ="
	objFile.WriteLine ""
	
	objFile.Close
	Set objFile = Nothing	
	Set objFSO = Nothing
	On Error Goto 0
End Sub

Sub LogError2File(strErrDesc)
	'This function will log the SQL errors to file.
	
	'If an error occurs at this stage, ignore it.
	On Error Resume Next		
	
	Dim objFSO	
	Dim strFileName
		
	Set objFSO = CreateObject("Scripting.FileSystemObject")	
	Set objFile = objFSO.OpenTextFile("EmailSQLErrors.txt", 8, True)
	
	objFile.WriteLine "Date : " & Now() 
	objFile.WriteLine strSQL
	objFile.WriteLine "SQLError : " & GetSQLErrors(ADOCn)
	objFile.WriteLine "Err.Description : " & strErrDesc
	objFile.WriteLine ""
	objFile.WriteLine "= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = ="
	objFile.WriteLine ""
	
	Set objFile = Nothing	
	Set objFSO = Nothing
	On Error Goto 0
End Sub

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
	WScript.Sleep 500	'Give the Output window a chance to render.
End Sub