   <#Backup RSSA dropbox, scheduled task (same name).
   #>
   Compress-Archive -Path 'C:\Users\Guy\Dropbox\RoySoc' -DestinationPath "D:\RSSA Backups\Dropbox\RSSA $(get-date -f yyyyMMdd-HHmmss).zip" -force
   