'Filename    : SendEmailFromTableWithSMTPMailer.vbs
'Created by  : Christo Pretorius on 8 July 2025
'Description : This script is used to send email from the Email table on the SQL server.

Dim blnGetEmails
Dim intEmails
Dim intSMTPMailers
Dim intLoop
Dim arrEmails
Dim arrSMTPInfo
Dim arrSMTP
Dim ADOCn
Dim strSQL
Dim strErrMsg
Dim objIEDebugWindow
Dim blnDebug
Dim blnLogging

blnDebug = True '###
blnLogging = True

'###
Call LogError2File("0", "Script started @ " & Now())

Call Main

Call LogError2File("0", "Script ended @ " & Now())
Call LogError2File("0", "= = = = = = = = = = = = = = = =")

wscript.quit 0 'Quit with success

Sub Main	
	Dim blnSMTPMailerIncluded
	Dim objSMTPMailer
	Dim strOutput
	Dim strReturnVal
	
	'Email values
	Dim intEmailID
	Dim strSubject
	Dim strSenderName
	Dim strSenderAddress
	Dim strRecipientName
	Dim strRecipientAddress
	Dim strBody
	Dim intPrecedence
	Dim blnIsHTML

	'SMTP fields
	Dim intWebsiteID
	Dim strServer
	Dim intPort
	Dim strUsername
	Dim strPassword
	Dim blnUseSecure
	Dim strReplyTo
	
	'Email errors
	Dim strResult1
	Dim strResult5
	Dim strResult9
	
	Call LogError2File("0", "Calling GetEmails()")
	
	Call GetEmails		
	'1st array fields:
	'(  
	' EmailID bigint,  
	' WebsiteID int,
	' Subject varchar(100),  
	' SenderName varchar(70),  
	' SenderAddress varchar(70),  
	' RecipientName varchar(70),  
	' RecipientAddress varchar(70),  
	' Body text, - 7500 chars
	' Precedence tinyint,  
	' PlainTextEmail tinyint  
	')  	
	
	'2nd array fields:
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
		
			If blnSMTPMailerIncluded = False Then	
				Call LogError2File("0", "Include clsSMTPMailer using ExecuteGlobal")
				
				'Include the contents of clsSMTPMailer.cls into this file and us it as if it is native code.
				On Error Resume Next
				Err.Clear
				ExecuteGlobal CreateObject("Scripting.FileSystemObject").OpenTextFile(".\clsSMTPMailer.vbs", 1).ReadAll
				blnSMTPMailerIncluded = True
				
				If Err.Number <> 0 Then
					Call LogError2File("0", "Including clsSMTPMailer.vbs failed. Err message: " & Err.Description)
					Call ResetAllEmails
					Exit Sub
				End If
								
				'Create an instance of the class
				Err.Clear
				Set objSMTPMailer = New clsSMTPMailer
				
				If Err.Number <> 0 Then
					Call LogError2File("0", "Create instance of clsSMTPMailer failed. Err message: " & Err.Description)
					Call ResetAllEmails
					Exit Sub
				End If
				
				Call LogError2File("0", "clsSMTPMailer included and initialised.")
			End If
					
			'Assign the email fields
			intEmailID = arrEmails(0 intLoop)
			intWebsiteID = arrEmails(1, intLoop)
			strSubject = arrEmails(2, intLoop)
			strSenderName = arrEmails(3, intLoop)
			strSenderAddress = LCase(arrEmails(4, intLoop))
			strRecipientName = Trim(arrEmails(5, intLoop))
			strRecipientAddress = arrEmails(6, intLoop)
			'strBody = arrEmails(7, intLoop)  Not assigned. It is up to 8000 characters. Let's save memory!
			intPrecedence = arrEmails(8, intLoop)
			
			blnIsHTML = True
			If arrEmails(9, intLoop) = 1 Then blnIsHTML = False
					
			'If the recipient address isn't empty...
			If Len(strRecipientAddress) > 0 Then 	
			
				'If it is not a dummy email...
				If strSenderAddress <> "dummy@dummy.dummy" Then
				
					'Find the sender's SMTP details
					intIndex = GetSMTPMailerIndex(intWebsiteID)
					
					If intIndex > -1 Then
						'Assign the SMTP fields
						strServer = arrSMTPInfo(1, intIndex)
						intPort = arrSMTPInfo(2, intIndex)
						strUsername = arrSMTPInfo(3, intIndex)
						strPassword = arrSMTPInfo(4, intIndex)
						blnUseSecure = arrSMTPInfo(5, intIndex)
						strReplyTo = arrSMTPInfo(6, intIndex)											
					Else
						'Send using postmaster@re4.co.za and local email server
						strServer = "196.22.138.229"
						intPort = 225
						strUsername = ""
						strPassword = ""
						blnUseSecure = False
						strSenderAddress = "postmaster@re4.co.za"
						strReplyTo = strSenderAddress			
					End If
					
					'Send the email using the SMTPMailer Class
					Call objSMTPMailer.SendSMTPMail ".\Emails\" & intEmailID, strServer, intPort, blnUseSecure, blnIsHTML, strUsername, strPassword, _
					  strSenderAddress, strSenderName, strReplyTo, strRecipientAddress, "", "", strSubject, arrEmails(7, intLoop), _
					  strOutput, strReturnVal
					  
					If strReturnVal = 0 Then
						'Success
						strResult1 = strResult1 & "," & intEmailID
					ElseIf strReturnVal = 1 Then
						'SMTP error
						strResult5 = strResult5 & "," & intEmailID
					Else
						'Other/unknown
						strResult9 = strResult95 & "," & intEmailID
					End If
					
				Else
					'Since it was a dummy email, flag it as successful.
					strResult1 = strResult1 & "," & intEmailID
				End If																
			Else
				strResult4 = strResult4 & "," & intEmailID
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
		arrEmails = ADORs.GetRows			'Put the email data in a 2 dimensional array
		intEmails = UBound(arrEmails, 2)	'Get the number of emails
		blnGetEmails = True
	End If
	
	Output intEmails & " emails to process"
	
	'###
	Call LogError2File("0", intEmails & " emails to process")
	
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
		Call LogSQLError2File(Err.Description)
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

Sub ResetAllEmails
	'This sub will update the retrieved emails statusses back to Unsent.	
	'This should only be called if no emails could be processed
	
	'Open a connection to the database.	
	On Error Resume Next	
	Err.Clear
	Set ADOCn = OpenADOConnection
		
	'Check if the connection opened successfully.
	If Err.number <> 0 Then
		On Error Goto 0
		Exit Function
	End If		
	
	strSQL = ""
	strIDs = ""

	For intLoop = 0 To intEmails
		strIDs = strIDs & "," & arrEmails(0, intLoop)
	Next
	
	'Remove the 1st comma
	strIDs = Mid(strIDs, 2)
	
	strSQL = "sp_UpdateEmailStatus '" & strIDs & "', 1"
	
	If blnDebug Then Output strSQL
	
	On Error Resume Next	
	Err.Clear 	
	Set ADORs = OpenADORsReadOnly(ADOCn, strSQL, True)
	
	If Err.number <> 0 Then
		Call LogSQLError2File(Err.Description)			
	End If
	
	On Error Goto 0
	Set ADORs = Nothing	'Disconnect recordset	
	Set ADOCn = Nothing	'Close the connection
End Sub

Function OpenADOConnection
	'This function will open a connection to the database
	'and return the connection.

	OpenADOConnection = "" 	'Assume failure

	Dim ADOCn
	Dim strConnString

	Set ADOCn = CreateObject("ADODB.Connection")
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

Sub LogError2File(strEmailID, strErrDesc)
	'This function will log the SMTP errors to file.
	
	If blnLogging = False Then Exit Sub
	
	'If an error occurs at this stage, ignore it.
	On Error Resume Next		
	
	Dim objFSO	
	Dim strFileName
		
	Set objFSO = CreateObject("Scripting.FileSystemObject")	
	Set objFile = objFSO.OpenTextFile("D:\ScheduledJobs\SendEmailFromTableWithSMTPMailer.log.txt", 8, True)
	
	If strEmailID <> "0" Then
		objFile.WriteLine "Date : " & Now() 	
		objFile.WriteLine "EmailID : " & strEmailID
		objFile.WriteLine "SMTP Error : " & strErrDesc		
	Else
		objFile.WriteLine strErrDesc		
	End If
	
	objFile.WriteLine ""
	
	objFile.Close
	Set objFile = Nothing	
	Set objFSO = Nothing
	On Error Goto 0
End Sub

Sub LogSQLError2File(strErrDesc)
	'This function will log the SQL errors to file.
	
	If blnLogging = False Then Exit Sub
	
	'If an error occurs at this stage, ignore it.
	On Error Resume Next		
	
	Dim objFSO	
	Dim objFile
	Dim strFileName
		
	Set objFSO = CreateObject("Scripting.FileSystemObject")	
	Set objFile = objFSO.OpenTextFile("D:\ScheduledJobs\SendEmailFromTableWithSMTPMailer_SQLErrors.txt", 8, True)
	
	objFile.WriteLine "Date : " & Now() 
	objFile.WriteLine strSQL
	objFile.WriteLine "SQLError : " & GetSQLErrors(ADOCn)
	objFile.WriteLine "Err.Description : " & strErrDesc
	objFile.WriteLine ""
	objFile.WriteLine "= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = ="
	objFile.WriteLine ""
	objFile.Close
	
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
