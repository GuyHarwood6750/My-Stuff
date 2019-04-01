$status = Test-Connection wserver
if ($status.statuscode -eq 0) {
    $file = Test-Path -path '\\wserver\kiosk\daily reports\preparation\Daily_snap*.xlsx'
    if ($file -eq $true) {
    $path = '\\wserver\Kiosk\Daily Reports\Preparation\'
    #$path = 'C:\Test\'
    $file = Get-ChildItem -Path $path -Name 'Daily_snapshot*.xlsx'
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
    $printarea = '$A$1:$K$' + $rows
    $xl.Sheets.Item('sheet1').PageSetup.Printarea = $printarea

   #$range4 = $xl.Range("I1").Entirecolumn
   #$range4.Select()
    #$range4.HorizontalAlignment = -4108

    #$range5 = $xl.Range("F1").Entirecolumn
    #$range5.Select()
    #$range5.HorizontalAlignment = -4108

    $range = $xl.Range("A:A").Entirecolumn
    $range.NumberFormat = 'yyyy/mm/dd'
    $range2 = $xl.Range("1:1").EntireRow
    $range2.Select()
    $range2.Font.Name = 'Calibri'
    $range2.Font.Bold = $true
    $range2.Font.ColorIndex = '-4105'
    $range3 = $xl.Range("K1").Entirecolumn
    #$range3.Select()
    #$xlfilter = "Arrived"
    #$range3.AutoFilter(11, $xlfilter)
    $rangefinal = $xl.Range("A1")
    $rangefinal.Select()

    $wb.save()
    $xl.Workbooks.Close()
    $xl.Quit()

    Get-process EXCEL | stop-process
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
   
    $src1 = '\\wserver\Kiosk\Daily Reports\'
    $dest1 = '\\wserver\kiosk\Daily Reports\old'
    
    Get-ChildItem -Path $src1\Daily_snapshot*.xlsx | Move-Item -Destination $dest1 -Force

   move-item -Path $a `
               -Destination '\\wserver\Kiosk\Daily Reports'

   Copy-Item -Path $src1\daily_snap*.xlsx -Destination 'C:\Userdata\Circe Launches\Daily Reports'          
   }
    else {Guy-SendGmail "Daily Snapshot file not found" "Check if script was run on server"}
}
 else {Guy-SendGmail "Connection to WSERVER does not exists" "PLEASE INVESTIGATE"}