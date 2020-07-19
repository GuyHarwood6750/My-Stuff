<# Warren Marine - 
    Report - Booking Voucher query
#>
$status = Test-Connection wserver
if ($status.statuscode -eq 0) {
    $file = Test-Path -path '\\wserver\wmarine\Booking Reports\BookingVCHQuery*.xlsx'
    if ($file -eq $true) { 
        $path = '\\wserver\wmarine\Booking Reports\'
        #$path = 'C:\Test\'
        $file = Get-ChildItem -Path $path -Name 'BookingVCHQuery*.xlsx'
        $insheet = $path + $file
       
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $xl.workbooks.Open($insheet)
        $xl.Sheets.Item('Sheet1').Activate()
    
        $xl.Sheets.Item('sheet1').PageSetup.Orientation = 2  # Landscape
        $xl.Sheets.Item('sheet1').PageSetup.Zoom = $false
        $Xl.Sheets.Item('sheet1').PageSetup.FitToPagesTall = 1
        $xl.Sheets.Item('sheet1').PageSetup.FitToPagesWide = 1
        $xl.Sheets.Item('sheet1').PageSetup.RightMargin = 0.89
        $xl.Sheets.Item('sheet1').PageSetup.LeftMargin = 15.02   #50.02

        $rows = $xl.Sheets.Item('sheet1').UsedRange.Rows.Count
        $printarea = '$A$1:$N$' + $rows
        $xl.Sheets.Item('sheet1').PageSetup.Printarea = $printarea

        $range5 = $xl.Range("I1").Entirecolumn
        $range5.Select()
        $range5.HorizontalAlignment = -4108

        $range5 = $xl.Range("F1").Entirecolumn
        $range5.Select()
        $range5.HorizontalAlignment = -4108

        $range5 = $xl.Range("E1").Entirecolumn
        $range5.Select()
        $range5.HorizontalAlignment = -4108

        $range5 = $xl.Range("J1").Entirecolumn
        $range5.Select()
        $range5.HorizontalAlignment = -4108

        $range = $xl.Range("B:B").Entirecolumn
        $range.NumberFormat = 'dd/mm/yyyy'
        $range2 = $xl.Range("1:1").EntireRow
        $range2.Select()
        $range2.Font.Name = 'Calibri'
        $range2.Font.Bold = $true
        $range2.Font.ColorIndex = '-4105'
                
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
        Get-ChildItem -Path $src2\BookingVCHQuery*.xlsx | Move-Item -Destination $dest2 -Force
      
        $dest1 = '\\wserver\wmarine\Booking Reports\Julia'
    
        Get-ChildItem -Path $path\BookingVCHQuery*.xlsx | Move-Item -Destination $dest1 -Force
 
        #Write-EventLog -LogName MyPowerShell -Source "WM" -EntryType Information -EventId 10 -Message "PrepaidCancelled script completed"

    }
    else {
        #Guy-SendGmail "Prepaid Cancelled script failed to copy to Server" "Check connection to Server"
        #Write-EventLog -LogName MyPowerShell -Source "WM" -EntryType Error -EventId 30 -Message "Prepaid Script failed, server not found"

    }
       
}       
else {
    $ThisScript = $MyInvocation.MyCommand.Name
    Guy-SendGmail "Connection to WServer does not exists!" "PLEASE INVESTIGATE - $ThisScript" 
    #Write-EventLog -LogName MyPowerShell -Source "GemTours" -EntryType Error -EventId 31 -Message "Script failed, VPN connection not found"

}