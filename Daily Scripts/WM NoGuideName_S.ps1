
    $file = Test-Path -path '\\wserver\wmarine\booking reports\julia\preparation\NoGuideName*.xlsx'
    if ($file -eq $true) {
        $path = '\\wserver\wmarine\booking reports\Julia\Preparation\'
        $file = Get-ChildItem -Path $path -Name 'NoGuideName*.xlsx'
        $a = $path + $file

        $src1 = '\\wserver\Kiosk\Daily Reports'
        $dest1 = '\\wserver\kiosk\Daily reports\old'
        $src2 = '\\wserver\WMarine\booking reports\Julia'
        $dest2 = '\\wserver\wmarine\booking reports\Julia\OLD'
    
        Get-ChildItem -Path $src1\NoGuideName*.xlsx | Move-Item -Destination $dest1 -Force
        Get-ChildItem -Path $src2\NoGuideName*.xlsx | Move-Item -Destination $dest2 -Force


        Copy-Item -path $a `
            -Destination '\\wserver\wmarine\kiosk\Daily Reports' 
        Move-Item -Path $a `
            -Destination '\\wserver\wmarine\booking reports\Julia'
            
        #Write-EventLog -LogName MyPowerShell -Source "WM" -EntryType Information -EventId 10 -Message "NoGuideName script completed"

    }
    Else { 
        #$ThisScript = $MyInvocation.MyCommand.Name
        #Guy-SendGmail "No Guide Name spreadsheet found" "Check if script ran on WSERVER - $ThisScript" 
    
        #Write-EventLog -LogName MyPowerShell -Source "WM" -EntryType Error -EventId 30 -Message "WM NoGuideName script failed, file not found on server"

    }
