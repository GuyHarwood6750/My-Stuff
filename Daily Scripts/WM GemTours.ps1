$status = Test-Connection wserver
if ($status.statuscode -eq 0) {
    $file = Test-Path -path '\\wserver\kiosk\invoices\Gem Tours\preparation\GemTours*.xlsx'
    if ($file -eq $true) { 
        $path = '\\wserver\Kiosk\invoices\Gem Tours\Preparation\'
        #$path = 'C:\Test\'
        $file = Get-ChildItem -Path $path -Name 'GemToursBookingDetails*.xlsx'
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
        $xl.Sheets.Item('sheet1').PageSetup.RightMargin = 0.89
        $xl.Sheets.Item('sheet1').PageSetup.LeftMargin = 50.02

        $rows = $xl.Sheets.Item('sheet1').UsedRange.Rows.Count
        $printarea = '$A$1:$L$' + $rows
        $xl.Sheets.Item('sheet1').PageSetup.Printarea = $printarea

        $range4 = $xl.Range("J1").Entirecolumn
        $range4.Select()
        $range4.HorizontalAlignment = -4108

        $range5 = $xl.Range("G1").Entirecolumn
        $range5.Select()
        $range5.HorizontalAlignment = -4108

        $range = $xl.Range("B:B").Entirecolumn
        $range.NumberFormat = 'dd/mm/yyyy'
        $range2 = $xl.Range("1:1").EntireRow
        $range2.Select()
        $range2.Font.Name = 'Calibri'
        $range2.Font.Bold = $true
        $range2.Font.ColorIndex = '-4105'
        $range3 = $xl.Range("K1").Entirecolumn
        $range3.Select()
        $xlfilter = "Arrived"
        $range3.AutoFilter(12, $xlfilter)
        $rangefinal = $xl.Range("A1")
        $rangefinal.Select()

        $wb.save()
        $xl.Workbooks.Close()
        $xl.Quit()

        Get-Process EXCEL | Stop-Process
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  
        $src1 = '\\wserver\Kiosk\Invoices\Gem Tours'
        $dest1 = '\\wserver\kiosk\Invoices\Gem Tours\OLD'
    
        Get-ChildItem -Path $src1\GemToursBookingDetails*.xlsx | Move-Item -Destination $dest1 -Force
 
        Move-Item -Path $a `
            -Destination '\\wserver\Kiosk\Invoices\Gem Tours'

        Write-EventLog -LogName MyPowerShell -Source "GemTours" -EntryType Information -EventId 10 -Message "GemTours script completed"

    }
    else { Guy-SendGmail "Gem Tours Booking details file not found" "Check if script was run on Server"
        Write-EventLog -LogName MyPowerShell -Source "GemTours" -EntryType Error -EventId 30 -Message "Script failed, file not found"

}
       
}       
else { Guy-SendGmail "Connection to WServer does not exists!" "WM GemTours" 
    Write-EventLog -LogName MyPowerShell -Source "GemTours" -EntryType Error -EventId 31 -Message "Script failed, VPN connection not found"

}