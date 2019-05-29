#Backup Warren Marine current financial year Pastel files.

Compress-Archive -Path "C:\Pastel18\CIRCE20A" -DestinationPath "D:\Pastel Backups - Hout Bay\2020\CIRCE20A $(get-date -f yyyyMMdd-HHmmss).zip" -force

<#
Compress-Archive -Path "C:\Pastel18\CIRCE20A" -DestinationPath "\\wserver\backup\PastelBKP\CIRCE20A $(get-date -f yyyyMMdd-HHmmss).zip" -force
#>
Write-EventLog -LogName MyPowerShell -Source "WM" -EntryType Information -EventId 10 -Message "Archive of CIRCE20A data completed"

