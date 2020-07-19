#Backup Warren Marine current financial year Pastel files.

Compress-Archive -Path "C:\Pastel18\CIRCE21A" -DestinationPath "D:\Pastel Backups - Hout Bay\2020\CIRCE21A $(get-date -f yyyyMMdd-HHmmss).zip" -force


Compress-Archive -Path "C:\Pastel18\CIRCE21A" -DestinationPath "\\wserver\backup\PastelBKP\CIRCE21A $(get-date -f yyyyMMdd-HHmmss).zip" -force

Write-EventLog -LogName MyPowerShell -Source "WM" -EntryType Information -EventId 10 -Message "Archive of CIRCE20A data completed"

