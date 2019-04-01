#Backup Warren Marine current financial year Pastel files.

    Compress-Archive -Path "C:\Pastel18\CIRCE20A" -DestinationPath "D:\Pastel Backups - Hout Bay\2020\circe20a $(get-date -f yyyyMMdd-HHmmss).zip" -force
