#Backup Route 62 Stop current financial year Pastel files.

Compress-Archive -Path "C:\Pastel18\CGC2019" -DestinationPath "D:\Pastel Backups\CGC\2019\CGC2019 $(get-date -f yyyyMMdd-HHmmss).zip" -force
<#
 Compress-Archive -Path "C:\Pastel18\CGC2019" -DestinationPath "\\wserver\backup\PastelBKP\R622020A $(get-date -f yyyyMMdd-HHmmss).zip" -force   
#>
#Write-EventLog -LogName MyPowerShell -Source "R62S" -EntryType Information -EventId 10 -Message "Archive of R622020A data completed"
