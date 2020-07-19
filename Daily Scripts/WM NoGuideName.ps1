$status = Test-Connection wserver
if ($status.statuscode -eq 0) {
    $file = Test-Path -path '\\wserver\wmarine\booking reports\julia\preparation\NoGuideName*.xlsx'
    if ($file -eq $true) {
        $path = '\\wserver\wmarine\booking reports\Julia\Preparation\'
        #$path = 'C:\Test\'
        $file = Get-ChildItem -Path $path -Name 'NoGuideName*.xlsx'
        $a = $path + $file

        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $xl.workbooks.Open($a)
        $ws = $xl.Sheets.Item('Sheet1').Activate()
    
        $xl.Sheets.Item('sheet1').PageSetup.Orientation = 2  # Landscape
        $xl.Sheets.Item('sheet1').PageSetup.Zoom = $false
        $Xl.Sheets.Item('sheet1').PageSetup.FitToPagesTall = 1
        $xl.Sheets.Item('sheet1').PageSetup.FitToPagesWide = 1
        $xl.Sheets.Item('sheet1').PageSetup.RightMargin = 10.89
        $xl.Sheets.Item('sheet1').PageSetup.LeftMargin = 40.02

        #$rows = $xl.Sheets.Item('sheet1').UsedRange.Rows.Count
        $printarea = '$A$1:$K$12'
        $xl.Sheets.Item('sheet1').PageSetup.Printarea = $printarea

        $range4 = $xl.Range("I1").Entirecolumn
        $range4.Select()
        $range4.HorizontalAlignment = -4108

        $range5 = $xl.Range("F1").Entirecolumn
        $range5.Select()
        $range5.HorizontalAlignment = -4108

        $range = $xl.Range("A:A").Entirecolumn
        $range.NumberFormat = 'yyyy/mm/dd'
        $range2 = $xl.Range("1:1").EntireRow
        $range2.Select()
        $range2.Font.Name = 'Calibri'
        $range2.Font.Bold = $true
        $range2.Font.ColorIndex = '-4105'

        $range2a = $xl.Range("2:22").EntireRow
        $range2a.Select()
        $range2a.Font.Name = 'Calibri'
        #$range2a.Font.Bold = $true
        $range2a.Font.ColorIndex = '3'

        $range3 = $xl.Range("K1").Entirecolumn
        $range3.Select()
        $xlfilter = "Unknown"
        $range3.AutoFilter(11, $xlfilter)
        $rangefinal = $xl.Range("A1")
        $rangefinal.Select()

        $wb.save()
        $xl.Workbooks.Close()
        $xl.Quit()

        #Get-Process EXCEL | Stop-Process
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)

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
            
        Write-EventLog -LogName MyPowerShell -Source "WM" -EntryType Information -EventId 10 -Message "NoGuideName script completed"

    }
    Else { 
        $ThisScript = $MyInvocation.MyCommand.Name
        Guy-SendGmail "No Guide Name spreadsheet found" "Check if script ran on WSERVER - $ThisScript" 
    
        Write-EventLog -LogName MyPowerShell -Source "WM" -EntryType Error -EventId 30 -Message "WM NoGuideName script failed, file not found on server"

    }
   }
 else {
    Guy-SendGmail "Connection to WSERVER does not exist" "PLEASE INVESTIGATE"

    Write-EventLog -LogName MyPowerShell -Source "WM" -EntryType Error -EventId 31 -Message "NoGuideName script failed, VPN connection not found"

 }