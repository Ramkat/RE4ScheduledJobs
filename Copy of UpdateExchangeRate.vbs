'Filename    : UpdateExchangeRate.vbs
'Created by  : Christo Pretorius 9 September 2005
'						 : Updated 26 Feb 2007
'Description : This script is used to retrieve the latest Rand/US Dollar exchange rate
'						 : from xe.com, update the exchange rate table and send an email.

Dim blnDebug
Dim sinExcRate

blnDebug = True '###
sinExcRate = ""

If GetExcRate Then	
	If UpdateExcRate Then
		Call EmailExcRate("Success")
	Else
		Call EmailExcRate("NoUpdate")
	End If	
Else
		Call EmailExcRate("Failed")
End If

WScript.Quit 0					'Quit with success

Function GetExcRate
	'This function will retrieve the exchange rate value from the xe.com website.
	
	GetExcRate = False
	
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
		
		If blnDebug Then wscript.echo "Cannot create XMLHTTP object" & vbcrlf & Err.Description
		Exit Function
	End If
		
	objXMLHTTP.Open "Get", "http://www.xe.com/ucc/convert.cgi?Amount=1&From=USD&To=ZAR", False
  objXMLHTTP.Send
  strResult = objXMLHTTP.ResponseText    
  Set objXMLHTTP = Nothing
  
  If blnDebug Then wscript.echo "Len(strResult) = " & vbcrlf & Len(strResult)
  
  If Len(strResult) = 0 Then				
		Exit Function
	End If 			
	
	'=====
	'Extract the exchange rate value.
	'=====		
	intPos1 = 0
	intPos1 = InStr(1, strResult, "ZAR,USD") 'Find the 1st ZAR,USD in the text	
	If intPos1 = 0 Then Exit Function
	
	intPos1 = InStr(intPos1 + 1, strResult, ">") 'Find the 1st > after the ZAR,USD in the text	
	If intPos1 = 0 Then Exit Function
	
	intPos2 = InStr(intPos1 + 1, strResult, "<") 'Find the 1st < after the > in the text	
	If intPos2 = 0 Then Exit Function

	'intPos2 = intPos1 - 8
	strValue = Trim(Mid(strResult, intPos1 + 2, intPos2 - intPos1 - 2)) 'Find the value between the > and <

If blnDebug Then
  wscript.echo "intPos 1 = " & intPos1 & " ; intPos2 = " & intPos2 & vbcrlf & "Value = " & strValue
End If
	
	If Not IsNumeric(strValue) Then Exit Function
	
	
	sinExcRate = CSng(strValue)
	sinExcRate = Round(sinExcRate, 2)
	
	'Ensure that a correct value was extracted.
	If sinExcRate < 1 Then Exit Function
	
	If blnDebug Then wscript.Echo "Exchange rate = " & sinExcRate
	GetExcRate = True
End Function

Function UpdateExcRate
	UpdateExcRate = False	'Assume failure
	
	Dim ADOCn
	Dim ADORs
	Dim strSQL
	
	strSQL = "sp_InsertExchangeRate 'USD'" _
					 & ", 'US Dollars'" _
					 & ", "  & sinExcRate  _
					 & ", "  & (1 / sinExcRate)
					 
	'Open a connection to the database.	
	On Error Resume Next
	Err.Clear
	Set ADOCn = OpenADOConnection
					 
	If Err.number <> 0 Then
		If blnDebug Then wscript.echo "Err = " & Err.Number
		Set ADOCn = Nothing
		On Error Goto 0
		Exit Function
	End If
	
	Set ADORs = OpenADORsReadOnly(ADOCn, strSQL, True)	
	
	If Err.number <> 0 Then
		Set ADOCn = Nothing
		Set ADORs = Nothing
		On Error Goto 0
		Exit Function
	End If
	
	Set ADOCn = Nothing
	Set ADORs = Nothing
	On Error Goto 0
	UpdateExcRate = True
End Function

Sub EmailExcRate(strAction)
	'This sub will email the exchange rate to info@thecampusgroup.com
	
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
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'Send the message using a network SMTP server.
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "127.0.0.1"
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 225 
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False 'Use SSL for the connection (True or False)
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "webmaster@re4.your-server.co.za"
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "TEFqzUZb"
	objSendMail.Configuration.Fields.Update
		
	objSendMail.From = chr(34) & "RE4 Exchange Rate Update" & chr(34) & " " & "info@thecampusgroup.com"
	objSendMail.To = "info@thecampusgroup.com"	
	objSendMail.TextBody = ""
	
	Select Case strAction
		Case "Success"
			objSendMail.Subject = "US$1 = R" & sinExcRate & " at " & Now()
			
		Case "Failed"
			objSendMail.Subject = "Unable to retrieve at " & Now()
			objSendMail.TextBody = "Manually update the exchage rate from" & vbcrlf & _
				"http://www.xe.com/ucc/convert.cgi?Amount=1&From=USD&To=ZAR"
			
		Case "NoUpdate"
			objSendMail.Subject = "Table update failed at " & Now()
			objSendMail.TextBody = "Manually update the exchage rate: US$1 = R" & sinExcRate
	End Select		
	
If blnDebug Then wscript.echo "Subject: " & objSendMail.Subject
	
	On Error Resume Next
	Err.Clear 
	objSendMail.Send
	
	If blnDebug And Err.Number <> 0 Then
		Wscript.echo "Email error: " & Err.Description
	End If
	
	Set objSendMail = Nothing	
	On Error Goto 0
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
	'This funciton will return a READ ONLY, ADO recordset.
	'If blnDisconnect = True, the recordset will be disconnected.

	OpenADORsReadOnly = "" 		'Assume failure

	Dim ADORs

If blnDebug Then wscript.echo "ADOConnection = " & TypeName(ADOConnection)
If blnDebug Then wscript.echo "SQL = " & strSQL

	Set ADORs = CreateObject("ADODB.Recordset")
If blnDebug Then wscript.echo "ADORs = " & TypeName(ADORs)

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