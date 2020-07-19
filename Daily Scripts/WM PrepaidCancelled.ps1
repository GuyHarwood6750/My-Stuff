<# Warren Marine - 
    Report - Prepaid group but trip was cancelled
#>
$status = Test-Connection wserver
if ($status.statuscode -eq 0) {
    $file = Test-Path -path 'C:\Userdata\Circe Launches\Daily Reports\Daily_Snapshot*.xlsx'
    if ($file -eq $true) { 
        $path = 'C:\Userdata\Circe Launches\Daily Reports\'
        #$path = 'C:\Test\'
        $file = Get-ChildItem -Path $path -Name 'Daily_Snapshot*.xlsx'
        $insheet = $path + $file
        $file2 = "Prepaid_Cancelled $(get-date -f yyyyMMdd-HHmm).xlsx"
        $outfile = $path + $file2

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
        $xl.Sheets.Item('sheet1').PageSetup.LeftMargin = 50.02

        $rows = $xl.Sheets.Item('sheet1').UsedRange.Rows.Count
        $printarea = '$A$1:$K$' + $rows
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
        $range3 = $xl.Range("K1").Entirecolumn
        $range3.Select()
        $xlfilter = "Cancel"
        $range3.AutoFilter(12, $xlfilter)
        $xlfilter = "PRPD"
        $range3.AutoFilter(7, $xlfilter)
                
        $rangefinal = $xl.Range("A1")
        $rangefinal.Select()
        
        $wb.saveas($outfile)
        $xl.Workbooks.Close()
        $xl.Quit()

        Get-Process EXCEL | Stop-Process
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
      
        $dest1 = '\\wserver\wmarine\Booking Reports\Julia'
    
        Get-ChildItem -Path $path\Prepaid_Cancelled*.xlsx | Move-Item -Destination $dest1 -Force
 
        Write-EventLog -LogName MyPowerShell -Source "WM" -EntryType Information -EventId 10 -Message "PrepaidCancelled script completed"

    }
    else {
        Guy-SendGmail "Prepaid Cancelled script failed to copy to Server" "Check connection to Server"
        Write-EventLog -LogName MyPowerShell -Source "WM" -EntryType Error -EventId 30 -Message "Prepaid Script failed, server not found"

    }
       
}       
else {
    $ThisScript = $MyInvocation.MyCommand.Name
    Guy-SendGmail "Connection to WServer does not exists!" "PLEASE INVESTIGATE - $ThisScript" 
    Write-EventLog -LogName MyPowerShell -Source "GemTours" -EntryType Error -EventId 31 -Message "Script failed, VPN connection not found"

}