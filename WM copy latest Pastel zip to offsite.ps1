#Warren Marine - Guy
#Copy the latest local backup file (zip) to offsite location.
#
#VPN connection must be present - need to write script to check!!
#
#New-PSDrive -Name "I" -PSProvider FileSystem -Root '\\wserver\WMarine\Finance\Pastel Offsite Backup\Pastel Backups\2019' -Persist
#  ----- use above line if you do not wish to keep a persistent drive ------ see command PSdrive ------
#
# Use ROBOCOPY as copy-item fails too often.
#
$vpnok = Get-VpnConnection
$locallconn = Test-NetConnection 'wserver' -InformationLevel Detailed

if ($vpnok.ConnectionStatus -eq 'connected')
{
        $DestinationFolder = "\\wserver\wmarine\Finance\Pastel Offsite Backup\Pastel Backups\2019"
        #$EarliestModifiedTime = (Get-date).AddDays(-1)
        $EarliestModifiedTime = (Get-Date).AddHours(-9)
        $srcefolder = "d:\pastel backups - hout bay\2019"
        $Files = Get-ChildItem $srcefolder\*.zip -File
        
        Remove-Item -path "$destinationfolder\*.zip" 

    foreach ($File in $Files) {
        if ($File.LastWriteTime -gt $EarliestModifiedTime)
     {
            robocopy $srcefolder $DestinationFolder $file.Name /v /is /log+:"c:\ps scripts logs\pastel.txt"
     }
    else 
    {
       #Write-Host "Not copying $File"
    }
}
}
    elseif ($locallconn.PingSucceeded -eq 'True') {
        
        $DestinationFolder = "\\wserver\wmarine\Finance\Pastel Offsite Backup\Pastel Backups\2019"
        #$EarliestModifiedTime = (Get-date).AddDays(-1)
        $EarliestModifiedTime = (Get-Date).AddHours(-9)
        $srcefolder = "d:\pastel backups - hout bay\2019"
        $Files = Get-ChildItem $srcefolder\*.zip -File
        
        Remove-Item -path "$destinationfolder\*.zip" 

    foreach ($File in $Files) {
        if ($File.LastWriteTime -gt $EarliestModifiedTime)
     {
            robocopy $srcefolder $DestinationFolder $file.Name /v /is /log+:"c:\ps scripts logs\pastel.txt"
     }
     else 
    {
       #Write-Host "Not copying $File"
    }
    }
}
    else 
    {
        Guy-SendGmail "Copy of Pastel zip file to offsite location failed" "PLEASE INVESTIGATE"
    }

