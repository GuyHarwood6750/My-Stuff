$credentials = Get-StoredCredential -AsCredentialObject -Target "GuyGmail"
$emailSmtpUser = $credentials.UserName
$emailSmtpPass = $credentials.Password
$emailSmtpServer = "smtp.gmail.com"
$emailSmtpServerPort = "587"
$emailFrom = "harwoodg123@gmail.com"
$emailTo = "accounts@harwood.co.za"
#$emailcc="myboss@gmail.com"

$emailMessage = New-Object System.Net.Mail.MailMessage( $emailFrom , $emailTo )
#$emailMessage.cc.add($emailcc)
$emailMessage.Subject = "My test mail" 
$emailMessage.Body = "Hello World"

$SMTPClient = New-Object System.Net.Mail.SmtpClient( $emailSmtpServer , $emailSmtpServerPort )
$SMTPClient.EnableSsl = $True
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential( $emailSmtpUser , $emailSmtpPass );
$SMTPClient.Send( $emailMessage )