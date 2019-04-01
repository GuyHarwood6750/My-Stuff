#Warren Marine - Guy Harwood
#Purpose: To copy the SQL database file from Server location to local USB drive & Dropbox.
#
#
    $vpnok = Get-VpnConnection
    $locallconn = Test-NetConnection -ComputerName "wserver.wmarine.local" -InformationLevel Detailed

if ($vpnok.ConnectionStatus -eq 'connected')
    {
            $sourcepath = "\\wserver\backup"
            $destpath = "D:\Circe Launches Backups\SQL Circe Bookings"
            $dropbox ="C:\Users\Guy\Dropbox\SCS\Circe Launches"

            Remove-Item -path "$destpath\circe*"
            Remove-item -Path "$dropbox\circe*"

        Get-ChildItem -path "$sourcepath\circe*" -file | 
            Sort-Object -Property Modifiedtime -Descending | Select-Object -First 1 | 

            copy-item -destination $destpath -Force
            copy-item -path "$destpath\circe*" -destination $dropbox -Force
     }
    elseif($locallconn.PingSucceeded -eq 'True'){
            $sourcepath = "\\wserver\backup"
            $destpath = "D:\Circe Launches Backups\SQL Circe Bookings"
            $dropbox ="C:\Users\Guy\Dropbox\SCS\Circe Launches"

            Remove-Item -path "$destpath\circe*"
            Remove-item -Path "$dropbox\circe*"

        Get-ChildItem -path "$sourcepath\circe*" -file | 
            Sort-Object -Property Modifiedtime -Descending | Select-Object -First 1 | 

            copy-item -destination $destpath -Force
            copy-item -path "$destpath\circe*" -destination $dropbox -Force

  }   
    
    else 
    {
        
        Guy-SendGmail "Copy of SQL backup file to USB drive & Dropbox failed" "PLEASE INVESTIGATE"
        
   }