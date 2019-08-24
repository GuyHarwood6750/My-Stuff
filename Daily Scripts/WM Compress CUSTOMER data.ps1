#Backup Warren Marine Customer files.
#
switch ((Get-date).DayOfWeek) {
    Monday {
        Compress-Archive -Path "\\wserver\wmarine\Customers\" -DestinationPath "d:\circe launches backups\Customers\Customers $(get-date -f yyyyMMdd-HHmmss).zip" -force
    }
    Tuesday {
        Compress-Archive -Path "\\wserver\wmarine\Customers\" -DestinationPath "d:\circe launches backups\Customers\Customers $(get-date -f yyyyMMdd-HHmmss).zip" -force
    }
    Wednesday {
        
    }
    Thursday {

    }
    Friday {
        $playsoung = New-Object System.Media.Soundplayer
        $playsoung.SoundLocation = 'C:\Users\Guy\Documents\Powershell\Sound\script start.WAV'
        $playsoung.playsync()

                
        $playsoung = New-Object System.Media.Soundplayer
        $playsoung.SoundLocation = 'C:\Users\Guy\Documents\Powershell\Sound\script end.WAV'
        $playsoung.playsync()

    }
    Saturday {
        
    }
    Sunday {

    }
    Default { }

}
[System.Console]::Beep(3000, 300)
[System.Console]::Beep(1000, 100)
[System.Console]::Beep(3000, 300)

 #Write-EventLog -LogName MyPowerShell -Source "WM" -EntryType Information -EventId 10 -Message "Backup of Customer data completed"

