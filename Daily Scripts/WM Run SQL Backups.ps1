switch ((get-date).Hour -le 14) {
    $true { . 'C:\Users\Guy\Documents\Powershell\Daily Scripts\WM copy latest sql backup file to local.ps1' }
    $false {
        . 'C:\Users\Guy\Documents\Powershell\Daily Scripts\WM copy latest sql backup file to local.ps1' 
        . 'C:\Users\Guy\Documents\Powershell\Daily Scripts\WM copy SQL backup file to NAS.ps1'
        }
        Default {
            . 'C:\Users\Guy\Documents\Powershell\Daily Scripts\WM copy latest sql backup file to local.ps1'
        }
    }
    