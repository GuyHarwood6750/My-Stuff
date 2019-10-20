#Backup Warren Marine Kiosk files.
#
switch ((Get-date).DayOfWeek) {
    Monday {
        Compress-Archive -Path "\\wserver\Kiosk\Invoices\*Scanned*" -DestinationPath "d:\circe launches backups\kiosk\Invoices\F2020\Scanned2020 $(get-date -f yyyyMMdd-HHmmss).zip" -force
    }
    Tuesday {

    }
    Wednesday {
        Compress-Archive -Path "\\wserver\Kiosk\Invoices\*Scanned*" -DestinationPath "d:\circe launches backups\kiosk\Invoices\F2020\Scanned2020 $(get-date -f yyyyMMdd-HHmmss).zip" -force
    }
    Thursday {

    }
    Friday {
        $playsoung = New-Object System.Media.Soundplayer
        $playsoung.SoundLocation = 'C:\Users\Guy\Documents\Powershell\Sound\script start.WAV'
        $playsoung.playsync()

        Compress-Archive -Path "\\wserver\Kiosk\Invoices\F2020" -DestinationPath "d:\circe launches backups\kiosk\Invoices\F2020\Invoices2020 $(get-date -f yyyyMMdd-HHmmss).zip" -force
        Compress-Archive -Path "\\wserver\Kiosk\Schedules" -DestinationPath "d:\circe launches backups\kiosk\Schedules\Schedules $(get-date -f yyyyMMdd-HHmmss).zip" -force
        
        $playsoung = New-Object System.Media.Soundplayer
        $playsoung.SoundLocation = 'C:\Users\Guy\Documents\Powershell\Sound\script end.WAV'
        $playsoung.playsync()

    }
    Saturday {
        Compress-Archive -Path "\\wserver\Kiosk\Invoices\*Scanned*" -DestinationPath "d:\circe launches backups\kiosk\Invoices\F2020\Scanned2020 $(get-date -f yyyyMMdd-HHmmss).zip" -force
    }
    Sunday {
        #Compress-Archive -Path "\\wserver\Kiosk\Invoices\F2020" -DestinationPath "d:\circe launches backups\kiosk\Invoices\F2020\Invoices2020 $(get-date -f yyyyMMdd-HHmmss).zip" -force
        #Compress-Archive -Path "\\wserver\Kiosk\Schedules" -DestinationPath "d:\circe launches backups\kiosk\Schedules\Schedules $(get-date -f yyyyMMdd-HHmmss).zip" -force
    }
    Default { }

}
[System.Console]::Beep(3000, 300)
[System.Console]::Beep(1000, 100)
[System.Console]::Beep(3000, 300)

 Write-EventLog -LogName MyPowerShell -Source "WM" -EntryType Information -EventId 10 -Message "Backup of KIOSK data completed"

