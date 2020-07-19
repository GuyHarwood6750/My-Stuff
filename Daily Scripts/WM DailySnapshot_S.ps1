    $file = Test-Path -path '\\wserver\wmarine\booking reports\Daily_snap*.xlsx'
    if ($file -eq $true) {
        $path = '\\wserver\wmarine\Booking Reports\'

        $file = Get-ChildItem -Path $path -Name 'Daily_snapshot*.xlsx'
        $a = $path + $file

        $src1 = '\\wserver\Kiosk\Daily Reports\'
        $dest1 = '\\wserver\kiosk\Daily Reports\old'

        $src2 = '\\wserver\wmarine\Booking Reports\julia\'
        $dest2 = '\\wserver\WMarine\booking reports\julia\old'
        
        #$src3 = 'C:\Userdata\Circe Launches\Daily Reports\'
        #$dest3 = 'C:\Userdata\Circe Launches\Daily Reports\old'

        Get-ChildItem -Path $src1\Daily_snapshot*.xlsx | Move-Item -Destination $dest1 -Force
        Get-ChildItem -Path $src2\Daily_snapshot*.xlsx | Move-Item -Destination $dest2 -Force
        #Get-ChildItem -Path $src3\Daily_snapshot*.xlsx | Move-Item -Destination $dest3 -Force


        Move-Item -Path $a `
            -Destination '\\wserver\Kiosk\Daily Reports'

        #Copy-Item -Path $src1\daily_snap*.xlsx -Destination 'C:\Userdata\Circe Launches\Daily Reports'  
        Copy-Item -Path $src1\daily_snap*.xlsx -Destination '\\wserver\WMarine\booking reports\julia' 
        
        #Write-EventLog -LogName MyPowerShell -Source "WM" -EntryType Information -EventId 10 -Message "DailySnapshot script completed"
        
    }
    else {
        #Guy-SendGmail "Daily Snapshot file not found" "Check if script was run on server"

        #Write-EventLog -LogName MyPowerShell -Source "WM" -EntryType Error -EventId 30 -Message "DailySnapshot script failed, file not found on server"

    }
