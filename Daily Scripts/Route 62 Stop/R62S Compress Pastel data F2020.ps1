#Backup Route 62 Stop current financial year Pastel files.

Compress-Archive -Path "C:\Pastel18\R622021A" -DestinationPath "D:\Pastel Backups\Route 62 Stop\2021\R622021A $(get-date -f yyyyMMdd-HHmmss).zip" -force
#
#Compress-Archive -Path "C:\Pastel18\R622020A" -DestinationPath "C:\USERS\GUY\DROPBOX\R62\ACCOUNTS\PASTEL BACKUP\R622020A $(get-date -f yyyyMMdd-HHmmss).zip" -force
<#
 Compress-Archive -Path "C:\Pastel18\R622020A" -DestinationPath "\\wserver\backup\PastelBKP\R622020A $(get-date -f yyyyMMdd-HHmmss).zip" -force   
#>
Write-EventLog -LogName MyPowerShell -Source "R62S" -EntryType Information -EventId 10 -Message "Archive of R622020A data completed"
