'Filename    : CheckAndRestartAppPools.vbs
'Created by  : christo.pretorius@gmail.com
'Created on  : 22 Apr 2025
'Description : Check if the app pools are running and restart those that arent'.

Dim shell, exec, output, line
Dim command, startCommand

' Create a shell object to run commands
Set shell = CreateObject("WScript.Shell")

' Command to list all application pools
command = "C:\Windows\System32\inetsrv\appcmd list apppool"

' Execute the command and capture the output
Set exec = shell.Exec(command)

' Read the output
output = ""
Do While exec.Status = 0
    WScript.Sleep 100
Loop
output = exec.StdOut.ReadAll()

'WScript.Echo "Application Pool list: " & output '############

' Split the output into lines
Dim lines
lines = Split(output, vbCrLf)

' Check each line for the application pool status
For Each line In lines
    If Trim(line) <> "" Then ' Ensure the line is not empty
        ' Check if the line contains the application pool name and status
        If InStr(line, "Stopped") > 0 Or InStr(line, "Unknown") > 0 Then
            ' Extract the application pool name from the line
            Dim appPoolName
            appPoolName = Split(line, """")(1) ' Get the app pool name			
            
            'WScript.Echo "Application Pool: " & appPoolName & " is stopped. Starting it..."
            ' Command to start the application pool
            startCommand = "C:\Windows\System32\inetsrv\appcmd start apppool /apppool.name:""" & appPoolName & """ /timeout:60000"
            ' Execute the start command
            shell.Run startCommand, 0, True  '0 = Hide the window. True = wait to complete before continueing
            'WScript.Echo "Application Pool: " & appPoolName & " has been started."
        Else
            ' Output the status of running application pools
            'WScript.Echo "Application Pool: " & line & " is running."
        End If
    End If
Next

' Clean up
Set exec = Nothing
Set shell = Nothing