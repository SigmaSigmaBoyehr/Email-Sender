param (
    [Parameter(Mandatory=$true)]
    [string]$subject
)

# Email settings
$EmailFrom = "" #write your email here
$SMTPServer = "smtp.gmail.com" // its smtp server of ur email
$SMTPPort = 587
$SMTPUsername = "" #write your email here (like emailfrom)
$SMTPPassword = "" #here is ur smtppassword

$delayBetweenEmails = 70 #delay in sec

# Meil recipients
$recipients = @(
    @{Name="Maxim";   Email="maxim@gmail.com"},
    @{Name="Артем";    Email="artyom@gmail.com"},
)


# Attache
$attachments = @(
    "C:\Users\...",
    "C:\Users\...",
)

# Body
$baseBody = @"
<p>Hello!<strong>I'm Vanya</strong>.</p>

<p>Here is ur text<strong>bold text</strong>text</p>

<p>Yours respectfully<br>
Ur name<br>
<strong>ur company</strong></p>
"@

try {
    #SMTPClient
    $SMTPClient = New-Object Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
    $SMTPClient.EnableSsl = $true
    $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($SMTPUsername, $SMTPPassword)

    foreach ($recipient in $recipients) {
        try {
            # Its for personalized mail
            $personalizedBody = @"
<html>
<head>
<style>
  body { font-family: Arial, sans-serif; line-height: 1.6; }
  strong { font-weight: bold; }
</style>
</head>
<body>
<p>$($recipient.Name), Hello!</p>
$baseBody
</body>
</html>
"@
            #Mail
            $mailMessage = New-Object Net.Mail.MailMessage
            $mailMessage.From = $EmailFrom
            $mailMessage.To.Add($recipient.Email)
            $mailMessage.Subject = $subject
            $mailMessage.Body = $personalizedBody
            $mailMessage.IsBodyHtml = $true

            #Attache
            foreach ($file in $attachments) {
                if (Test-Path $file) {
                    $attachment = New-Object Net.Mail.Attachment($file)
                    $mailMessage.Attachments.Add($attachment)
                    Write-Host "Добавлено вложение: $file" -ForegroundColor Cyan
                }
            }

            #Send mail
            Write-Host "Отправка письма для $($recipient.Name) <$($recipient.Email)>..." -ForegroundColor Yellow
            $SMTPClient.Send($mailMessage)
            Write-Host "Успешно отправлено!" -ForegroundColor Green
            
            #CLear
            if ($mailMessage.Attachments) {
                $mailMessage.Attachments.Dispose()
            }
	    if ($recipient -ne $recipients[-1]) {
                Write-Host "Ожидание $delayBetweenEmails секунд перед следующей отправкой..." -ForegroundColor Gray
                Start-Sleep -Seconds $delayBetweenEmails
            }
        }
        catch {
            Write-Host "Ожидание $delayBetweenEmails секунд перед следующей отправкой..." -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
        }
    }
}
catch {
    Write-Host "Общая ошибка SMTP:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    
    if ($_.Exception.Message -match "5.7.0") {
        Write-Host @"
        
!ТРЕБУЕТСЯ ДОПОЛНИТЕЛЬНАЯ НАСТРОЙКА GMAIL:

1. Проверьте пароль приложения
2. Разрешите доступ для ненадежных приложений
3. Разблокируйте аккаунт по ссылке:
   https://accounts.google.com/DisplayUnlockCaptcha
"@ -ForegroundColor Yellow
    }
}
finally {
    if ($SMTPClient) {
        $SMTPClient.Dispose()
    }
}
