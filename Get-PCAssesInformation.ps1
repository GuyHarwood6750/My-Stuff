$computersystem = (Get-WmiObject win32_computersystem)
$bios = (Get-WmiObject win32_bios)
$results = @{'Computer Name' = $computersystem.Name;
            'Model' = $computersystem.model;
            'Serial Number' = $bios.SerialNumber}
$Report = New-Object -TypeName psobject -Property $results
Clear-Host
$report