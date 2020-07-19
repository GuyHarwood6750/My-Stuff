   <#Backup R62S dropbox, scheduled task (same name).
   #>
   Compress-Archive -Path 'C:\Users\Guy\Dropbox\R62' -DestinationPath "D:\R62S DROPBOX BACKUPS\R62S $(get-date -f yyyyMMdd-HHmmss).zip" -force
   
    #Write-EventLog -LogName MyPowerShell -Source "R62S" -EntryType Information -EventId 10 -Message "Archive of R62S Dropbox data completed"

   # Write-EventLog -LogName MyPowerShell -Source "R62S" -EntryType Information -EventId 10 -Message "Archive of R62S2019 Pastel data completed"