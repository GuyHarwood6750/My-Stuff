   <#Backup my Powershell scripts to offisite location, scheduled task (same name).
   #>
   Compress-Archive -Path 'C:\Users\Guy\Documents\Powershell' -DestinationPath "\\wserver\software\scripts backup\guy\Powershell $(get-date -f yyyyMMdd-HHmmss).zip" -force

   Compress-Archive -Path 'C:\Users\Guy\Documents\Route 62 Stop' -DestinationPath "\\wserver\software\scripts backup\guy\route62stop $(get-date -f yyyyMMdd-HHmmss).zip" -force
    
   #Write-EventLog -LogName MyPowerShell -Source "Guy" -EntryType Information -EventId 10 -Message "Archive of Powershell scripts completed"

    Compress-Archive -Path 'C:\users\guy\documents\WindowsPowerShell' -DestinationPath "\\wserver\software\scripts backup\guy\WindowsPowershell $(get-date -f yyyyMMdd-HHmmss).zip" -force
    
    #Write-EventLog -LogName MyPowerShell -Source "Guy" -EntryType Information -EventId 10 -Message "Archive of WindowsPowershell scripts completed"