#$date = Get-Date
#$day = Get-Date -Format 'ddddd'
#$time = (Get-Date).ToString("hh:mm")
    Switch ((get-date).ToString('tt')) {
        'AM' {"Morning script"}
        'PM' {"Afternoon script"}
        Default {Write-Output "time is not what it seems"}
    }

    switch ((get-date).Hour -le 10) {
        $true {. 'C:\Users\Guy\Documents\Powershell\Daily Scripts\copy latest sql backup file to local3.ps1'}
        $false {" after 10"}
        Default {}
    }