REM Malware bytes windows firewall control
net start wfcs

REM Windows defender firewall
net start mpssvc

net start "AVG Antivirus"
net start "AVG Tools"
net start MSSQL$SQL2017
net start SQLAgent$SQL2017
net start SQLWriter

REM MailEnable Mail Transfer Agent
net start MEMTAS
REM MailEnable SMTP Connection
net start MESMTPCS

REM World Wide Web Publishing Service
net start W3SVC
