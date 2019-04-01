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

        $range2a = $xl.Range("2:12").EntireRow
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

        Get-process EXCEL | stop-process
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)

        $src1 = '\\wserver\Kiosk\'
        $dest1 = '\\wserver\kiosk\old'
    
        Get-ChildItem -Path $src1\NoGuideName*.xlsx | Move-Item -Destination $dest1 -Force

        Copy-item -path $a `
            -Destination '\\wserver\wmarine\kiosk' 
        Move-item -Path $a `
            -Destination '\\wserver\wmarine\booking reports\Julia'
    }
    Else {Guy-SendGmail "No Guide Name spreadsheet found" "Check if script ran on WSERVER"}
   }
 else {
    Guy-SendGmail "Connection to WSERVER does not exist" "PLEASE INVESTIGATE"
 }