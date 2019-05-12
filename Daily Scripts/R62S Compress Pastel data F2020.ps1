#Backup Route 62 Stop current financial year Pastel files.

    Compress-Archive -Path "C:\Pastel18\R622020A" -DestinationPath "D:\Pastel Backups\Route 62 Stop\2020\R622020A $(get-date -f yyyyMMdd-HHmmss).zip" -force
    Write-EventLog -LogName MyPowerShell -Source "R62S" -EntryType Information -EventId 10 -Message "Archive of R622020A data completed"
