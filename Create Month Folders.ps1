   
    $listfolders = Import-csv -path 'C:\Userdata\Circe Launches\Monthly Invoices\Folder months.csv' | Where-Object {$_.Ind -eq 1}
        Foreach ($listf in $listfolders) {
            [int]$day = $($listf.days)
            $month = $($listf.Month)
            $Year = $($listf.Year)
            $scan = $($listf.Scan)
            $doctype = $($listf.DocType)
      
    $monthfolder = "$scan"+" $month"+" $doctype"+" $Year"
    #$mainloc = 'c:\test\monthly folders\'+"$monthfolder"+"\"
    $mainloc = '\\wserver\wmarine\kiosk\schedules\F2021\'+"$monthfolder"+"\"
    
     Do
     { 
     #$finalfolder = "$mainloc"+" $month"+" $year"               #create month only
     $finalfolder = "$mainloc"+"$day"+" $month"+" $year"       #create month & each day of month
     New-Item -Path $finalfolder -ItemType 'directory' -force
     #$day
     $day--
     } until ($day -le 0)
    }