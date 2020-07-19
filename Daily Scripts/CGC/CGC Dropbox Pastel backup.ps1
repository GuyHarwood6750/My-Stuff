   <#Backup CGC dropbox and Pastel data, scheduled task (same name).
   #>
   Compress-Archive -Path 'C:\Users\Guy\Dropbox\CGC' -DestinationPath "D:\CGC DROPBOX BACKUPS\CGC $(get-date -f yyyyMMdd-HHmmss).zip" -force
    Write-EventLog -LogName MyPowerShell -Source "CGC" -EntryType Information -EventId 10 -Message "Archive of CGC Dropbox data completed"

    Compress-Archive -Path 'C:\PASTEL18\CGC2019' -DestinationPath "D:\Pastel Backups\CGC\2019\CGC2019 $(get-date -f yyyyMMdd-HHmmss).zip" -force
    Write-EventLog -LogName MyPowerShell -Source "CGC" -EntryType Information -EventId 10 -Message "Archive of CGC2019 Pastel data completed"