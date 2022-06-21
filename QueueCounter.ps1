# Logging
New-Item -itemType Directory -Path "C:\" -Name "Temp" -ErrorAction SilentlyContinue
New-Item -ItemType File -Path "C:\Temp" -Name "MailQueue.log" -Force
Start-Transcript -Append C:\Temp\MailQueue.log

# Server and Email variables
$MailCount = (Get-ChildItem C:\inetpub\mailroot\queue).Count
$SMTPServer = "mail.server.com"
$EmailSubject = "Mail Count Has Reached Threshold"
$EmailBody = "Mail count on " + $SMTPServer + " is " + $MailCount

if ($MailCount -ge 100) {
    Send-MailMessage -From email@domain.com -To email@domain.com -Subject $EmailSubject -Body $EmailBody -SmtpServer $SMTPServer
    Write-Output $MailCount
    Stop-Transcript
}
else {
    Write-Output $MailCount
    Stop-Transcript
}
