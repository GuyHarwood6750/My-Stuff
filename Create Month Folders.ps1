   
    $listfolders = Import-csv -path 'C:\Userdata\Circe Launches\Monthly Invoices\Folder months.csv' | Where-Object {$_.Ind -eq 1}
        Foreach ($listf in $listfolders) {
            [int]$day = $($listf.days)
            $month = $($listf.Month)
            $Year = $($listf.Year)
            $scan = $($listf.Scan)
            $doctype = $($listf.DocType)
      
    $monthfolder = "$scan"+" $month"+" $doctype"+" $Year"
    $mainloc = 'c:\test\Fuel Invoices\'+"$monthfolder"+"\"
    #$mainloc = '\\wserver\wmarine\kiosk\invoices\'+"$monthfolder"+"\"
    
     Do
     { 
     $finalfolder = "$mainloc"+"$day"+" $month"+" $year"
     New-Item -Path $finalfolder -ItemType 'directory' -force
     #$day
     $day--
     } until ($day -le 0)
    }