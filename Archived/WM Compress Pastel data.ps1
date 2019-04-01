$status = Test-Connection wserver
 if ($status.statuscode -eq 0) {

    #Remove-Item -Path "D:\Pastel Backups - Hout Bay\2020\circe20a*.zip"
    #Remove-Item -Path "\\wserver\wmarine\finance\Pastel Offsite Backup\Pastel Backups\2020\circe20a*.zip"

    Compress-Archive -Path "C:\Pastel18\CIRCE20A" -DestinationPath "D:\Pastel Backups - Hout Bay\2020\circe20a $(get-date -f yyyyMMdd-HHmmss).zip" -force

    #robocopy "D:\Pastel Backups - Hout Bay\2020" '\\wserver\wmarine\Finance\Pastel Offsite Backup\Pastel Backups\2020' 'circe20a*.zip' /v /is /log+:"c:\ps scripts logs\pastel2.txt"
        #robocopy "D:\Pastel Backups - Hout Bay\2019" 'C:\Test' 'circe19a*.zip' /v /is /log+:"c:\ps scripts logs\pastel2.txt"
 }
else {
    Guy-SendGmail "Pastel backup failed." "PLEASE INVESTIGATE"
 }