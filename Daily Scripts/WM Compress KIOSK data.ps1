#Backup Warren Marine Kiosk files.
#$date = Get-Date
#$day = (Get-Date).DayOfWeek
switch ((Get-date).DayOfWeek) {
    Monday {
        Compress-Archive -Path "\\wserver\Kiosk\Invoices\F2020" -DestinationPath "d:\circe launches backups\kiosk\Invoices\F2020\Invoices2020 $(get-date -f yyyyMMdd-HHmmss).zip" -force
    }
    Tuesday {

    }
    Wednesday {

    }
    Thursday {

    }
    Friday {

    }
    Saturday {
        Compress-Archive -Path "\\wserver\Kiosk\Invoices\*Scanned*" -DestinationPath "d:\circe launches backups\kiosk\Invoices\F2020\Scanned2020 $(get-date -f yyyyMMdd-HHmmss).zip" -force
    }
    Sunday {
        Compress-Archive -Path "\\wserver\Kiosk\Invoices\*Scanned*" -DestinationPath "d:\circe launches backups\kiosk\Invoices\F2020\Scanned2020 $(get-date -f yyyyMMdd-HHmmss).zip" -force
    }
    Default { }
}





# Write-EventLog -LogName MyPowerShell -Source "WM" -EntryType Information -EventId 10 -Message "Backup of KIOSK data completed"

