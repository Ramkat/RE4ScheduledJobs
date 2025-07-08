C:\Windows\System32\inetsrv\appcmd stop apppool /apppool.name:"secure.re4.co.za" /timeout:60000
C:\Windows\system32\inetsrv\appcmd start apppool /apppool.name:"secure.re4.co.za" /timeout:60000
C:\Windows\system32\inetsrv\appcmd stop apppool /apppool.name:"DefaultAppPool" /timeout:60000
C:\Windows\system32\inetsrv\appcmd start apppool /apppool.name:"DefaultAppPool" /timeout:60000
C:\Windows\system32\inetsrv\appcmd stop apppool /apppool.name:"MailEnableAppPool" /timeout:60000
C:\Windows\system32\inetsrv\appcmd start apppool /apppool.name:"MailEnableAppPool" /timeout:60000
C:\Windows\system32\inetsrv\appcmd stop apppool /apppool.name:"MailEnableEASPool" /timeout:60000
C:\Windows\system32\inetsrv\appcmd start apppool /apppool.name:"MailEnableEASPool"