    switch ((get-date).Hour -le 13) {
        $true {. 'C:\Users\Guy\Documents\Powershell\Daily Scripts\copy latest sql backup file to local3.ps1'}
        $false {
           . 'C:\Users\Guy\Documents\Powershell\Daily Scripts\copy latest sql backup file to local3.ps1' 
           . 'C:\Users\Guy\Documents\Powershell\Daily Scripts\copy SQL backup file to NAS.ps1'
        }
        Default {}
    }
    