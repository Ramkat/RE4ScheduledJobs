'Filename    : SendEmailFromTableWithSMTPMailer.vbs
'Created by  : Christo Pretorius on 8 July 2025
'Description : This script is used to send email from the Email table on the SQL server.

Option Explicit

'####
'To update for LIVE server:
'
'OpenADOConnection -> connection String
'Replace C:\\Webs\\secure.re4.co.za\\   with C:\\Webs\\RE4\\secure.re4.co.za\\
'Replace C:\Webs\ScheduledJobs\  with  D:\ScheduledJobs\
'Comment out lines that contain '### DELETE
'Ensure blnDebug = False
'Ensure blnLogging = False

Dim blnGetEmails
Dim intEmails
Dim intSMTPMailers
Dim intLoop
Dim arrEmails
Dim arrSMTPInfo
Dim arrSMTP
Dim ADOCn
Dim ADORs
Dim strSQL
Dim strErrMsg
Dim objIEDebugWindow
Dim blnDebug
Dim blnLogging

'EmailIDs for different EmailResultIDs
Dim strResult1
Dim strResult5
Dim strResult9

blnDebug = False '###
blnLogging = False

Call LogError2File("0", "Script started @ " & Now())

Call Main

Call LogError2File("0", "Script ended @ " & Now())
Call LogError2File("0", "= = = = = = = = = = = = = = = =" & vbcrlf)

wscript.quit 0 'Quit with success

Sub Main		
	Dim objSMTPMailer
	Dim strOutput
	Dim strReturnVal
	Dim blnHaveDefSMTPDetails
	Dim blnCanSend
	Dim intIndex
	
	'Email values
	Dim intEmailID
	Dim strSubject
	Dim strSenderName
	Dim strSenderAddress
	Dim strRecipientName
	Dim strRecipientAddress
	Dim strBody	
	Dim blnIsHTML

	'SMTP fields
	Dim intWebsiteID
	Dim strServer
	Dim intPort
	Dim strUsername
	Dim strPassword
	Dim blnUseSecure
	Dim strReplyTo	
	
	'Try to include the SMTPMailer's wrapper class into this script.
	Call LogError2File("0", "Include clsSMTPMailer using ExecuteGlobal")
					
	On Error Resume Next
	Err.Clear
	
	'Include the contents of file clsSMTPMailer.vbs into this one.
	ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\\Webs\\secure.re4.co.za\\SecureRE4Scripts\\SMTPMailer\\clsSMTPMailer.vbs", 1).ReadAll		
	
	If Err.Number <> 0 Then
		Call LogError2File("0", "Including file clsSMTPMailer.vbs failed. Err message: " & Err.Description)		
		Exit Sub
	End If
					
	'Create an instance of the class
	Err.Clear
	Set objSMTPMailer = New clsSMTPMailer
	
	If Err.Number <> 0 Then
		Call LogError2File("0", "Create instance of clsSMTPMailer failed. Err message: " & Err.Description)		
		Exit Sub
	End If
	
	Call LogError2File("0", "clsSMTPMailer included and initialised.")
	
	'Get the default SMTP details.
	ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\\Webs\\secure.re4.co.za\\SecureRE4Scripts\\SMTPMailer\\SMTPDefaults.vbs", 1).ReadAll	
		
	'If no error occurred and we have SMTP details...
	If Err.Number = 0 And conSMTPServer <> "" And conSMTPPort <> "" Then
		blnHaveDefSMTPDetails = True
	End If
	
	Call LogError2File("0", "Calling GetEmails()")
	
	Call GetEmails		
	'1st array's fields:
	'(  
	' EmailID bigint,  
	' WebsiteID int,
	' Subject varchar(100),  
	' SenderName varchar(70),  
	' SenderAddress varchar(70),  
	' RecipientName varchar(70),  
	' RecipientAddress varchar(70),  
	' PlainTextEmail tinyint  
	' Body text, - 7500 chars	
	')  	
	
	'2nd array's fields:
	'(  
	' WebsiteID
	' SMTPServer
	' SMTPPort
	' SMTPUser
	' SMTPPassword
	' SMTPSecure (True/False)
	' EmailReplyTo
	')  	
		
	Do While blnGetEmails											
		'Loop through the email array.	
		For intLoop = 0 To intEmails								
			blnCanSend = False
			
			'Assign the email fields.
			intEmailID = arrEmails(0, intLoop)
			intWebsiteID = arrEmails(1, intLoop)
			strSubject = arrEmails(2, intLoop)
			strSenderName = arrEmails(3, intLoop)
			strSenderAddress = LCase(arrEmails(4, intLoop))
			strRecipientName = Trim(arrEmails(5, intLoop))
			strRecipientAddress = arrEmails(6, intLoop)
			'strBody = arrEmails(8, intLoop)  Not assigned. It is up to 8000 characters. Let's save memory! Read direct from the array.
			
			blnIsHTML = True
			If arrEmails(7, intLoop) = 1 Then blnIsHTML = False
					
			'If the recipient address isn't empty...
			If Len(strRecipientAddress) > 0 Then 	
			
				'If it is not a dummy email...
				If strSenderAddress <> "dummy@dummy.dummy" Then
				
					'Find the sender's SMTP details.
					intIndex = GetSMTPMailerIndex(intWebsiteID)
					
					If intIndex > -1 Then
						'Assign the SMTP fields.
						strServer = arrSMTPInfo(1, intIndex)
						intPort = arrSMTPInfo(2, intIndex)
						strUsername = arrSMTPInfo(3, intIndex)
						strPassword = arrSMTPInfo(4, intIndex)
						blnUseSecure = arrSMTPInfo(5, intIndex)
						strReplyTo = arrSMTPInfo(6, intIndex)
						blnCanSend = True
						
					ElseIf blnHaveDefSMTPDetails Then
						'Send using the default SMTP details.
						strServer = conSMTPServer
						intPort = conSMTPPort
						strUsername = conSMTPUser
						strPassword = conSMTPPassword
						blnUseSecure = conSMTPSecure
						strSenderAddress = conSenderAddress
						strReplyTo = conSenderAddress
						blnCanSend = True
					End If
					
					If blnCanSend Then
						On Error Resume Next
						Err.Clear
						
						'Send the email using the SMTPMailer Class.
						objSMTPMailer.SendSMTPMail "C:\Webs\secure.re4.co.za\SecureRE4Scripts\SMTPMailer\Emails\" & intEmailID, strServer, intPort, blnUseSecure, blnIsHTML, strUsername, strPassword, _
						  strSenderAddress, strSenderName, strReplyTo, strRecipientAddress, "", "", strSubject, blnLogging, arrEmails(8, intLoop), _
						  strOutput, strReturnVal
						  
						If Err.Number > 0 Then
							Call LogError2File(intEmailID, "Error calling objSMTPMailer.SendSMTPMail: " & Err.Description)
							strReturnVal = 9
						End If
						
						Call LogError2File(intEmailID, "Email sent") '###
						 
						If strReturnVal = 0 Then
							'Success
							strResult1 = strResult1 & "," & intEmailID
						ElseIf strReturnVal = 1 Then
							'SMTP error
							strResult5 = strResult5 & "," & intEmailID
						Else
							'Other/unknown
							strResult9 = strResult9 & "," & intEmailID
						End If					
					Else
						'We do not have SMTP details.
						'Other/unknown
						strResult9 = strResult9 & "," & intEmailID
					End If
				Else
					'Since it was a dummy email, flag it as successful.
					strResult1 = strResult1 & "," & intEmailID
				End If																
			Else
				'Other/unknown
				strResult9 = strResult9 & "," & intEmailID
			End If
		Next
		
		If UpdateResults Then
			Call GetEmails
		Else
			blnGetEmails = False
		End If
	Loop
	
	Set objSMTPMailer = Nothing		
End Sub

Sub GetEmails
	'This sub will retrieve unsent emails.
	blnGetEmails = False
	intEmails = 0
	strSQL = "sp_GetEmailToSendWithSmtpInfo"	'Note: This SP returns 2 recordsets.
	
	'Open a connection to the database.	
	On Error Resume Next
	Err.Clear
	Set ADOCn = OpenADOConnection

	'Check if the connection opened successfully.
	If Err.number <> 0 Then		
		Call LogError2File("0", "OpenADOConnection didn't execute in GetEmails()" & vbcrlf & Err.Description)
		Call LogError2File("0", GetSQLError())
		Exit Sub
	End If
	
	'Call LogError2File("0", "ADOConnection opened...") '### Comment out!
	
	Err.Clear 
	Set ADORs = OpenADORsReadOnly(ADOCn, strSQL, False)
	
	If Err.number <> 0 Then
		Call LogError2File("0", "OpenADORsReadOnly didn't execute in GetEmails()" & vbcrlf & Err.Description)
		Call LogError2File("0", GetSQLError())
		Set ADORs = Nothing
		Set ADOCn = Nothing
		Exit Sub
	End If
	
	'Call LogError2File("0", "ADORs opened...") '### Comment out!
			
	If ADORs.RecordCount > 0 Then
		arrEmails = ADORs.GetRows			'Put the email data in a 2 dimensional array
		intEmails = UBound(arrEmails, 2)	'Get the number of emails
		blnGetEmails = True
	End If
	
	Output (intEmails + 1) & " emails to process"
	
	'###
	Call LogError2File("0", (intEmails + 1) & " emails to process")
	
	'Get the next recordset - it contains the SMTP details.
	Set ADORs = ADORs.NextRecordset
	
	If ADORs.RecordCount > 0 Then
		arrSMTPInfo = ADORs.GetRows				'Put the SMTP data in a 2 dimensional array				
		intSMTPMailers =  UBound(arrEmails, 2)	'Get the number of SMTPMailers
	End If	

	Set ADORs = Nothing
	Set ADOCn = Nothing
End Sub

Function GetSMTPMailerIndex(intWebsiteID)
	'This function will loop through arrSMTPInfo and return the record number in the array
	'that contains the sender's address
	Dim intSMTPLoop
	
	For intSMTPLoop = 0 To intSMTPMailers
		If arrSMTPInfo(0, intSMTPLoop) = intWebsiteID Then
			GetSMTPMailerIndex = intSMTPLoop
			Exit Function
		End If
	Next
	
	GetSMTPMailerIndex = -1	'Error - sender wasn't found!
End Function

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
		Call LogError2File("0", "OpenADOConnection didn't execute in UpdateResults()" & vbcrlf & Err.Description)
		Call LogError2File("0", GetSQLError())
		Exit Function
	End If		
		
	If strResult1 <> "" Then
		'Remove the 1st comma
		strResult1 = Mid(strResult1, 2)
		strSQL = "exec sp_UpdateEmailResult '" & strResult1 & "', 1 ; "
	End If
		
	If strResult5 <> "" Then
		'Remove the 1st comma
		strResult5 = Mid(strResult5, 2)
		strSQL = strSQL & "exec sp_UpdateEmailResult '" & strResult5 & "', 5 ; "
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
		Call LogError2File("0", "OpenADORsReadOnly didn't execute in UpdateResults()" & vbcrlf & Err.Description)
		Call LogError2File("0", GetSQLError())
		Set ADORs = Nothing
		Set ADOCn = Nothing
		On Error Goto 0
		Exit Function
	End If
	
	On Error Goto 0
	Set ADORs = Nothing	'Disconnect recordset
	strResult1 = ""
	strResult5 = ""	
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
	strConnString = "Driver={SQL Server Native Client 11.0};pwd=~letzgetem@1l!;uid=EmailUser;database=wlc;Server=WINSRV2025STD\MSSQLSERVER2017"	

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

Sub LogError2File(strEmailID, strErrDesc)
	'This function will log the SMTP errors to file.
	
	If blnLogging = False Then Exit Sub
	
	'If an error occurs at this stage, ignore it.
	On Error Resume Next		
		
	Dim objFSO	
	Dim objFile	
		
	Set objFSO = CreateObject("Scripting.FileSystemObject")	
	Set objFile = objFSO.OpenTextFile("C:\Webs\ScheduledJobs\SendEmailFromTableWithSMTPMailer.log.txt", 8, True) '8 = For appending
	
	If strEmailID <> "0" Then
		objFile.WriteLine "Date : " & Now() 	
		objFile.WriteLine "EmailID : " & strEmailID
		objFile.WriteLine strErrDesc & vbcrlf
	Else
		objFile.WriteLine strErrDesc		
	End If	
	
	objFile.Close
	Set objFile = Nothing	
	Set objFSO = Nothing
	On Error Goto 0
End Sub

Function GetSQLError()
	'This function will return the SQL statement and errors
		
	GetSQLError = "SQL: " & strSQL & vbcrlf & "SQLError: " & GetSQLErrors(ADOCn)	 
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

Sub Output(myText)
	If blnDebug = False Then Exit Sub	'Ensure that if we don't show unnecessary windows if debugging isn't enabled.

	If Not IsObject(objIEDebugWindow) Then
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
