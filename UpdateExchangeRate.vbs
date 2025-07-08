'Filename    : UpdateExchangeRate.vbs
'Created by  : Christo Pretorius 9 September 2005
'			 : Updated 26 Feb 2007
'			 : Updated 24 Aug 2017 to use resbank.co.za
'			 : Updated 8 Dec 2020 to use x-rates.com
'Description : This script is used to retrieve the latest Rand/US Dollar exchange rate
'			 : from a web site, update the exchange rate table and send an email.


SetLocale(1033)  'US English - needed to ensure that numeric functions uses a dot for the decimal separator.

Dim blnDebug
Dim blnLog
Dim sinExcRate
Dim objIEDebugWindow

blnDebug = False '###
blnLog = False '###
sinExcRate = ""

Log "=== UpdateExchangeRate started at " & Now() & " ==="

If GetExcRate Then	
	Log "GetExcRate = True"
	If UpdateExcRate Then		
		Log "UpdateExcRate = True"
		Call EmailExcRate("Success")
	Else
		Log "UpdateExcRate = False"
		Call EmailExcRate("NoUpdate")
	End If	
Else
	Log "GetExcRate = False"
	Call EmailExcRate("Failed")
End If

Log "~~~ UpdateExchangeRate ended ~~~" & vbCrlf

WScript.Quit 0					'Quit with success

Function GetExcRate
	'This function will retrieve the exchange rate value from a web site.
	
	GetExcRate = False
	
	'If blnDebug Then 
		Log "Starting GetExcRate"		
		Output "Starting GetExcRate"		
	'End If
	
	Dim objXMLHTTP
	Dim strResult
	Dim intPos1
	Dim intPos2
	Dim strValue
    
	On Error Resume Next
	Err.Clear 
	Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")

	If Err.number <> 0 Then  		
		Set objXMLHTTP = Nothing
		
		Log "Cannot create XMLHTTP object" & vbcrlf & Err.Description			
		Output "Cannot create XMLHTTP object" & vbcrlf & Err.Description			
		Exit Function
	End If
	
	Log "objXMLHTTP created"
		
	objXMLHTTP.Open "Get", "https://www.x-rates.com/calculator/?from=USD&to=ZAR&amount=1", False
	objXMLHTTP.Send
	strResult = objXMLHTTP.ResponseText    
	Set objXMLHTTP = Nothing

	'If blnDebug Then 
		Log "Len(strResult) = " & Len(strResult)
		Output "Len(strResult) = " & Len(strResult) & vbcrlf & "<textarea cols=80 rows=20>" & strResult & "</textarea>"
		'Output strResult
	'End If

	If Len(strResult) = 0 Then				
		Exit Function
	End If 				
	
	'=====
	'Extract the exchange rate value.
	'=====			
	intPos1 = 0
	intPos1 = InStr(1, strResult, "ccOutputRslt") 'Find the 1st R/$ text
	
	'If blnDebug Then 	
		Output "intPos1 = " & intPos1	
	'End If
	
	If intPos1 = 0 Then Exit Function	
	strValue = Trim(Mid(strResult, intPos1 + 14, 5)) 'Find the numeric value		
	
	'If blnDebug Then
		  'Output "intPos 1 = " & intPos1 & " ; intPos2 = " & intPos2 & vbcrlf & "Value = " & strValue
		  Output Mid(strResult, intPos1, 30)
	'End If
	
	If IsValidNumeric(strValue, False, False, True) = False Then
		'If blnDebug Then 
			Log "IsValidNumeric(strValue) = False !!!"
			Output "IsValidNumeric(strValue) = False !!!"
		'End If
		
		Exit Function
	End If		
		
	Log "strValue = " & strValue
	
	'For some strange reason, when this script is run from a task scheduler, CSng() doesn't execute correctly.
	'Thus we are going to bypass it using the exec statement
	'sinExcRate = CSng(strValue)
	Dim sExec
	sExec = "sinExcRate = CSng(" & strValue & ")"
	Execute sExec
	
	Log "CSng(strValue) = " & sinExcRate
	Output "CSng(strValue) = " & sinExcRate	
	
	sinExcRate = Round(sinExcRate, 2)
	Log "Round(sinExcRate, 2) = " & sinExcRate
	Output "Round(sinExcRate, 2) = " & sinExcRate
	
	'Ensure that a correct value was extracted.
	If sinExcRate < 1 Then 
		Log "sinExcRate < 1 (" & sinExcRate & ")."
		Output "sinExcRate < 1 (" & sinExcRate & ")."
		Exit Function
	End If
	
	
	
	'If blnDebug Then 
		Log "Finishing GetExcRate. Exchange rate = " & sinExcRate
		Output "Finishing GetExcRate. Exchange rate = " & sinExcRate
	'End If
	
	GetExcRate = True
End Function

Function UpdateExcRate
	UpdateExcRate = False	'Assume failure
	
	Log "Starting UpdateExcRate"
	Output "Starting UpdateExcRate"
	
	Dim ADOCn
	Dim ADORs
	Dim strSQL
	
	Dim sExec
	Dim sExcRate
	
	On Error Resume Next
	Err.Clear
	sExec = "sExcRate = 1 / " & sinExcRate	
	Log "sExec = " & sExec
		
	Execute sExec	
	Log "sExcRate = " & sExcRate
	Log "TypeName(sExcRate) = " & TypeName(sExcRate)
		
	sExcRate = Round(sExcRate, 2)
	
	strSQL = "sp_InsertExchangeRate 'USD'" _
					 & ", 'US Dollars'" _
					 & ", "  & sinExcRate  _
					 & ", "  & sExcRate
					 
	Log strSQL
					 
	'Open a connection to the database.	
	On Error Resume Next
	Err.Clear
	Set ADOCn = OpenADOConnection
	
	Log "ADOCn = " & Typename(ADOCn)
	Log "strSQL = " & strSQL
	Output "ADOCn = " & Typename(ADOCn)
	Output "strSQL = " & strSQL
					 
	If Err.number <> 0 Then
		'If blnDebug Then 
			Log "Create connection Err = " & Err.Number
			Output "Create connection Err = " & Err.Number
		'End If
				
		Set ADOCn = Nothing
		On Error Goto 0
		Exit Function
	End If
	
	Set ADORs = OpenADORsReadOnly(ADOCn, strSQL, True)	
	
	If Err.number <> 0 Then
		Log "Open ADORs Err = " & Err.Number
		Output "Open ADORs Err = " & Err.Number
			
		Set ADOCn = Nothing
		Set ADORs = Nothing
		On Error Goto 0
		Exit Function
	End If
	
	Set ADOCn = Nothing
	Set ADORs = Nothing
	On Error Goto 0
	UpdateExcRate = True
	
	If blnDebug Then Output "Finishing UpdateExcRate"
End Function

Sub EmailExcRate(strAction)
	'This sub will email the exchange rate to info@thecampusgroup.com
	
	Log "Starting EmailExcRate(" & strAction & ")"
	Output "Starting EmailExcRate(" & strAction & ")"
	
	Dim objSendMail	
	
	On Error Resume Next
	Err.Clear
	'Initialise the SMTP mailer object
	Set objSendMail = CreateObject("CDO.Message")
		
	If Err.Number <> 0 Then
		Log "No email object created." & vbcrlf & Err.Description
		Output "No email object created." & vbcrlf & Err.Description
		
		Set objSendMail = Nothing					
		Exit Sub
	End If	

	Log "TypeName(objSendMail) = " & TypeName(objSendMail)
				
	On Error Goto 0
	Err.Clear	
	
	'Set a few values regarding the SMTP server, port etc.
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'Send the message using a network SMTP server.
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "cluster1out.eu.messagelabs.com"
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False 'Use SSL for the connection (True or False)
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "info@wildlifecampus.com"
	'objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "W0J!s74X"
	'objSendMail.Configuration.Fields.Update
	
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'Send the message using a network SMTP server.
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "196.22.138.229"
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 225 
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False 'Use SSL for the connection (True or False)
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
'	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "webmaster@re4.your-server.co.za"
'	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "TEFqzUZb"
	objSendMail.Configuration.Fields.Update
		
	objSendMail.From = chr(34) & "RE4 Exchange Rate Update" & chr(34) & " " & "info@thecampusgroup.com"
	objSendMail.To = "info@thecampusgroup.com"	
	'objSendMail.To = "christo.pretorius@gmail.com"	
	objSendMail.TextBody = ""
	
	Select Case strAction
		Case "Success"
			objSendMail.Subject = "US$1 = R" & sinExcRate & " at " & Now()
			
		Case "Failed"
			objSendMail.Subject = "Unable to retrieve at " & Now()
			objSendMail.TextBody = "Manually update the exchange rate from" & vbcrlf & _
				"https://www.x-rates.com/calculator/?from=USD&to=ZAR&amount=1"
			
		Case "NoUpdate"
			objSendMail.Subject = "Table update failed at " & Now()
			objSendMail.TextBody = "Manually update the exchange rate: US$1 = R" & sinExcRate
	End Select		
	
	Output "Subject: " & objSendMail.Subject
	Output "Calling objSendMail.Send (to " & objSendMail.To & ")"
	Log "Calling objSendMail.Send (to " & objSendMail.To & ")"
	
	On Error Resume Next
	Err.Clear 
		
	If blnDebug = False Then 		
		objSendMail.Send
	End If
	
	If Err.Number <> 0 Then
		Log "Email send error: " & Err.Description
		Output "Email send error: " & Err.Description
	End If
	
	'Also send the email to Eloise. Sharon requested to disable this on 15 May 2022.
	'objSendMail.To = "eloise@wildlifecampus.com"
	'objSendMail.To = "christo.pretorius@gmail.com"	
	'
	'On Error Resume Next
	'Err.Clear 
	'	
	'If blnDebug = False Then 
	'	Log "Calling objSendMail.Send (to " & objSendMail.To & ")"
	'	objSendMail.Send
	'End If
	'
	'If Err.Number <> 0 Then
	'	Log "Email send error: " & Err.Description
	'	Output "Email send error: " & Err.Description
	'End If
	'
	'Set objSendMail = Nothing	
	'On Error Goto 0
	
	Output "Finishing EmailExcRate"
End Sub

Function OpenADOConnection
	'This function will open a connection to the database
	'and return the connection.

	OpenADOConnection = "" 	'Assume failure

	Dim ADOCn
	Dim strConnString

	Set ADOCn = CreateObject("ADODB.Connection")
	
	'ODBC Connection
	strConnString = "driver={SQL Server Native Client 11.0};pwd=~letzgetle@rn1ng!;uid=WlcUser;database=wlc;Server=127.0.0.1\SQL2017"
	'Azure strConnString = "driver={SQL Server Native Client 11.0};pwd=~letzgetle@rn1ng!;uid=WlcUser;database=wlc;Server=20.50.111.103"
	'###strConnString = "driver={SQL Server};pwd=letzgetle@rn1ng;uid=wlcuser;database=wlc;Server=10.5.200.16"
	
	On Error Resume Next

	ADOCn.CommandTimeout = 60
	ADOCn.CursorLocation = 3	'Client-side cursor. DO NOT CHANGE IT !!!
	ADOCn.Open strConnString

	If Not ADOCn Is Nothing Then
		Set OpenADOConnection = ADOCn
	End If
	
	On Error Goto 0
End Function

Function OpenADORsReadOnly(ADOConnection, strSQL, blnDisconnect)
	'This function will return a READ ONLY, ADO recordset.
	'If blnDisconnect = True, the recordset will be disconnected.

	OpenADORsReadOnly = "" 		'Assume failure

	Dim ADORs

	If blnDebug Then Output "ADOConnection = " & TypeName(ADOConnection)
	If blnDebug Then Output "SQL = " & strSQL

	Set ADORs = CreateObject("ADODB.Recordset")
	If blnDebug Then Output "ADORs = " & TypeName(ADORs)

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

Function IsValidNumeric(ByVal strNumericValue, blnAllowMinus, blnAllowPlus, blnAllowDecPoint)
	'This function is used to check if a value contains only numerics and if
	'the user wants, + - . and , signs. 
		
	'Note: The IsNumeric() function in VBS allow the letters R/r thus a
	'value like 33r or 33R is a valid numeric value.

	'Note: Commas are not allowed. If you save 1,1 into a SQL table,
	'it changes the value to 11.
	
	
	IsValidNumeric = False		'Assume failure
	
	Dim intCount	
			
	'Remove leading and trailing spaces.
	strNumericValue = Trim(strNumericValue)

	If Len(strNumericValue) = 0 Then
		Exit Function
	End If
	
	'Check first character for + and - and digits
	Select Case Left(strNumericValue, 1)
		Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
		Case "-"
			If Not blnAllowMinus Then
				Exit Function
			End If
				
		Case "+"
			If Not blnAllowPlus Then
				Exit Function
			End If
			
		Case "."
			If Not blnAllowDecPoint Then
				Exit Function
			Else
				'Decimal point may occur only once.
				blnAllowDecPoint = false
			End If

		Case Else
			Exit Function
	End Select
			
	'Loop through the 2nd to 2nd last character.
	If Len(strNumericValue) > 2 Then
		For intCount = 2 to Len(strNumericValue) - 1
			Select Case Mid(strNumericValue, intCount, 1)
				Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
				Case "."
					If Not blnAllowDecPoint Then
						Exit Function
					Else
						'May occur only once
						blnAllowDecPoint = False
					End If								
		
				Case Else				
					Exit Function
			End Select
		Next
	End If
	
	'Check last character for + and - and digits
	Select Case Right(strNumericValue, 1)
		Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
		Case "-"
			If Not blnAllowMinus Then
				Exit Function
			End If
				
		Case "+"
			If Not blnAllowPlus Then
				Exit Function
			End If				

		Case Else
			Exit Function
	End Select
	
	'The value checks out fine.
	IsValidNumeric = True		
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
End Sub

Sub Log(msg)
	If blnLog = False Then Exit Sub 
		
	Dim objFSO
	Dim objFile	
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile("D:\ScheduledJobs\UpdateExchangeRate.log.txt", 8, True)	'8=For Appending
	
	objFile.WriteLine msg	
	objFile.Close
	
	Set objFile = Nothing
	Set objFSO = Nothing
End Sub