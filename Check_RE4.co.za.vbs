'Filename    : Check_RE4.co.za.vbs
'Created by  : Christo Pretorius 4 Jan 2010
'Description : This script is used to retrieve the homepage of re4.co.za. If it fails, it will
'							 update a counter in a text file. If the counter reaches 2, then it will send an
'							 email to administrators and reset the counter to 0.

Dim blnDebug
Dim intDownCount
Dim dteNow

dteNow = Now()
blnDebug = False
intDownCount = 0

If GetWebsite Then	
	Call UpdateCounter(0)
Else
	Call UpdateCounter(1)
End If

WScript.Quit 0					'Quit with success

Function GetWebsite
	'This function will retrieve the homepage of Re4.co.za and search for a string at the bottom of the page.
	
	GetWebsite = False
	
	Dim objXMLHTTP
  Dim strResult
  Dim intPos1
  Dim strValue
    
  On Error Resume Next
  Err.Clear 
  Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
  
  If Err.number <> 0 Then  		
		Set objXMLHTTP = Nothing
		
		If blnDebug Then wscript.echo "Cannot create XMLHTTP object" & vbcrlf & Err.Description
		Exit Function
	End If
		
	objXMLHTTP.Open "Get", "http://www.re4.co.za", False
  objXMLHTTP.Send
  strResult = objXMLHTTP.ResponseText    
  Set objXMLHTTP = Nothing
  
  If blnDebug Then wscript.echo "Len(strResult) = " & vbcrlf & Len(strResult)
  
  If Len(strResult) = 0 Then	Exit Function
	
	'=====
	'Search for Information Act
	'=====		
	intPos1 = 0
	intPos1 = InStr(1, strResult, "Information Act") 
	If intPos1 = 0 Then Exit Function
		
	If blnDebug Then
	  wscript.echo "'Information Act' at position " & intPos1
	End If
	
	GetWebsite = True
End Function

Sub UpdateCounter(intVal)
	Dim objFSO
	Dim objFile
	Dim strLine
		
	'Create an object of the file system.
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	On Error Resume Next
	
	If intVal = 0 Then
		'The website is up, so ensure that the counter is 0 by deleting the file and creating it.
		objFSO.DeleteFile("Check_RE4.co.za.log")
		Set objFile = objFSO.OpenTextFile("Check_RE4.co.za.log", 8, True)
		objFile.WriteLine "0"
	Else
		'The website is down. Check the status of the previous attempt to test the website.
		Set objFile = objFSO.OpenTextFile("Check_RE4.co.za.log", 1, True)
		strLine = Trim(objFile.ReadLine)
		objFile.Close
		
		If blnDebug Then wscript.echo "strLine = " & strLine
			
		If Len(strLine) > 0 Then 
			strLine = CInt(strLine)
		Else
			strLine = 0
		End If
			
		If strLine = 2 Then 
			'The counter is already 2, which means this is the 3rd time the website is down.
			'Thus reset the counter to 0 (by deleting and creating the file) and send an email to the administrators.
			objFSO.DeleteFile("Check_RE4.co.za.log")
			Set objFile = objFSO.OpenTextFile("Check_RE4.co.za.log", 8, True)
			objFile.WriteLine "0"

			On Error Goto 0			
			Call EmailError
		Else
			objFSO.DeleteFile("Check_RE4.co.za.log")
			Set objFile = objFSO.OpenTextFile("Check_RE4.co.za.log", 8, True)
			objFile.WriteLine(CInt(strLine) + 1)
		End If
	End If
	
	objFile.Close
	Set objFile = Nothing
	Set objFSO = Nothing
	On Error Goto 0
End Sub

Sub EmailError
	'This sub will send an email to the administrators.
	
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
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1 'Send the message using the local SMTP server.
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False 'Use SSL for the connection (True or False)
	objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	objSendMail.Configuration.Fields.Update
		
	objSendMail.From = chr(34) & "RE4 Web Server" & chr(34) & " " & "info@thecampusgroup.com"
	objSendMail.To = "christo@thecampusgroup.com"	
	objSendMail.Subject = "Server Monitor - RE4.co.za"
	objSendMail.TextBody = "RE4.co.za is down on " & GetWeekday & " " & Day(dteNow) _
		& " " & MonthName(Month(dteNow)) & " " & Year(dteNow) & " at " & Hour(dteNow) _
		& ":" & Minute(dteNow) & ":" & Second(dteNow)
	
	On Error Resume Next
	Err.Clear 
	objSendMail.Send
	
	objSendMail.To = "info@wildlifecampus.com"	
	objSendMail.Send
	
	'objSendMail.To = "0828570995@voda.co.za"	
	'objSendMail.Send
	
	If blnDebug And Err.Number <> 0 Then
		Wscript.echo "Email error: " & Err.Description
	End If
	
	Set objSendMail = Nothing	
	On Error Goto 0
End Sub

Function GetWeekday
	Select Case Weekday(dteNow)
		Case vbSunday			GetWeekday = "Sunday"
		Case vbMonday			GetWeekday = "Monday"
		Case vbTuesday		GetWeekday = "Tuesday"
		Case vbWednesday	GetWeekday = "Wednesday"
		Case vbThursday		GetWeekday = "Thursday"
		Case vbFriday			GetWeekday = "Friday"
		Case vbSaturday		GetWeekday = "Saturday"
	End Select
End Function