Option Explicit
Dim objShell
Dim lngReturnValue

Set objShell = CreateObject("WScript.Shell")
lngReturnValue = objShell.Run("wscript.exe C:\ScheduledJobs\BackupWebsitesToFTP.vbs 20200819", 7, False)

wscript.echo "return val = " & lngReturnValue
Set objShell = Nothing