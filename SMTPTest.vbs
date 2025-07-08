'Filename    : SMTPTest.vbs
'Created by  : Christo Pretorius 4 March 2014
'Description : This script is used to test the SMTP server (hMailServer)

Dim blnDebug
Dim sinExcRate

blnDebug = True '###
Call Test
WScript.Quit 0					'Quit with success


Sub Test
	Dim objSendMail
	
If blnDebug = False Then
	wscript.echo "Starting email sending"
	On Error Resume Next
	Err.Clear
End IF

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
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'Send the message using a network SMTP server.
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "196.22.138.229"
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 225 
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False 'Use SSL for the connection (True or False)
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
'	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "EmailUser" ' "webmaster@re4.your-server.co.za"
'	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "~letzgetem@1l!" '"TEFqzUZb"
	objSendMail.Configuration.Fields.Update
		
	objSendMail.From = chr(34) & "RE4 web server" & chr(34) & " " & "christo.pretorius@gmail.com"
	objSendMail.To = "investorcampus@re4.co.za"
	objSendMail.Subject = "SMTP Test from NEW web server"	
	objSendMail.TextBody = "Email test from RE4's NEW web server at " & Now() & "." & vbcrlf & "Eloise, when you get this, please let me know."
	
	
	If blnDebug = False Then
		On Error Resume Next
		Err.Clear 
	End If

	objSendMail.Send
	
	If blnDebug Then
		If Err.Number <> 0 Then
			Wscript.echo "Email error: " & Err.Description
		End If

		wscript.echo "Email send complete."
	End If
	
	Set objSendMail = Nothing	
	On Error Goto 0
End Sub