$status = Test-Connection wserver
if ($status.statuscode -eq 0) {
    $file = Test-Path -path '\\wserver\wmarine\booking reports\Julia\Preparation\BookingRateisZero*.xlsx'
    if ($file -eq $true) {
        $path = '\\wserver\wmarine\booking reports\Julia\Preparation\'
        #$path = 'C:\Test\'
        $file = Get-ChildItem -Path $path -Name 'BookingRateisZero*.xlsx'
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

        $range5 = $xl.Range("E1").Entirecolumn
        $range5.Select()
        $range5.HorizontalAlignment = -4108

        $range = $xl.Range("B:B").Entirecolumn
        $range.NumberFormat = 'dd/mm/yyyy'
        $range2 = $xl.Range("1:1").EntireRow
        $range2.Select()
        $range2.Font.Name = 'Calibri'
        $range2.Font.Bold = $true
        $range2.Font.ColorIndex = '-4105'

        $range2a = $xl.Range("K:K").Entirecolumn
        #$range2a = $xl.Range("1:1").EntireRow
        $range2a.Select()
        $range2a.Font.Name = 'Calibri'
        $range2a.Font.Bold = $true
        $range2a.Font.ColorIndex = '3'
               
        $rangefinal = $xl.Range("A1")
        $rangefinal.Select()

        $wb.save()
        $xl.Workbooks.Close()
        $xl.Quit()

        #Get-Process EXCEL | Stop-Process
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)

        #Move the file from previous day to 'OLD"    
        $src2 = '\\wserver\WMarine\booking reports\Julia'
        $dest2 = '\\wserver\wmarine\booking reports\Julia\OLD'
        Get-ChildItem -Path $src2\BookingRateisZero*.xlsx | Move-Item -Destination $dest2 -Force

        #Move today's file to Current location
        Move-Item -Path $a `
            -Destination '\\wserver\wmarine\booking reports\Julia'
            
        Write-EventLog -LogName MyPowerShell -Source "WM" -EntryType Information -EventId 10 -Message "BookingRateZero script completed"

    }
    Else { Guy-SendGmail "BookingRateZero spreadsheet not found" "Check if script ran on WSERVER - Script -> BookRateZero" 
    
        Write-EventLog -LogName MyPowerShell -Source "WM" -EntryType Error -EventId 30 -Message "BookingRateZero script failed, file not found on server"

}
   }
 else {
    Guy-SendGmail "Connection to WSERVER does not exist" "BookingRateZero Script - failed"

    Write-EventLog -LogName MyPowerShell -Source "WM" -EntryType Error -EventId 31 -Message "BookingRateZero script failed, VPN connection not found"

 }