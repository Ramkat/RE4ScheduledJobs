'Filename    : SendEmailTest.vbs

Dim blnDebug
Dim sinExcRate

blnDebug = True '###

Call SendEmailTest

wscript.echo "Email test complete"
WScript.Quit 0					'Quit with success



Sub SendEmailTest
		
	Dim objSendMail
	
	On Error Resume Next
	Err.Clear
	'Initialise the SMTP mailer object
	Set objSendMail = CreateObject("CDO.Message")
		
	If Err.Number <> 0 Then
		If blnDebug Then Wscript.echo "No email object created." & vbcrlf & Err.Description
		
		Set objSendMail = Nothing					
		Exit Sub
	End If		
				
	On Error Goto 0
	Err.Clear	
	
	'Set a few values regarding the SMTP server, port etc.
	'Sending using hMailServer via re4.your-server.co.za
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'Send the message using a network SMTP server.
	''objSendMailSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "196.22.138.230"
	''objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 225 '587
	''objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False 'Use SSL for the connection (True or False)
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
	''objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
		
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "ScheduledJobs"
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "~Sch3du1edJ0bs!"	'This password is set in hMailServer!
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "webmaster@re4.your-server.co.za"	
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "S@f@r1S@f@r1"	'This password is set in hMailServer!
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "info@wildlifecampus.com"
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "W0J!s74X"	
	
	'New POP account created by Dewald @ Global Micro - 13 Dec 2021.
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'Send the message using a network SMTP server.
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp3.dotnetwork2.co.za"
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False 'Use SSL for the connection (True or False)
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "webservices@anchorcapital.co.za"
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "dd25X3BG"
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	
	objSendMail.Configuration.Fields.Update
		
	objSendMail.From = chr(34) & "Christo" & chr(34) & " " & "christo@thecampusgroup.com"
	objSendMail.To = "christo.pretorius@gmail.com"	
	objSendMail.TextBody = "christo@Campus to Christo@gmail via webservices@anchor 9:28"
	objSendMail.Subject = "christo@Campus to Christo@gmail via webservices@anchor 9:28"
	
	If blnDebug Then 
		wscript.echo "Subject: " & objSendMail.Subject
	End If
	
	On Error Resume Next
	Err.Clear 
	objSendMail.Send
	
	If blnDebug And Err.Number <> 0 Then
		Wscript.echo "Email error: " & Err.Description
	End If
	
	Set objSendMail = Nothing	
	On Error Goto 0
End Sub
