$sourcepath = "\\wserver\backup"
$destpath = "D:\Circe Launches Backups\SQL Circe Bookings"

Remove-Item -path "$destpath\circe*"

Get-ChildItem -path "$sourcepath\circe*" -file | 
    Sort-Object -Property Modifiedtime -Descending | Select-Object -First 1 | 
    copy-item -destination $destpath -Force