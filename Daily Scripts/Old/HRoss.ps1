$status = Test-Connection wserver
 if ($status.statuscode -eq 0) {
    $path = '\\wserver\Kiosk\invoices\Hylton Ross\Preparation\'
    $file = Get-ChildItem -Path $path -Name 'HrossBookingDetails*.xlsx'
    $a = $path + $file

    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $wb = $xl.workbooks.Open($a)
    $ws = $xl.Sheets.Item('Sheet1').Activate()
    $range = $xl.Range("A:A").Entirecolumn
    $range.NumberFormat = 'yyyy/mm/dd'
    $range2 = $xl.Range("1:1").EntireRow
    $range2.Select()
    $range2.Font.Name = 'Calibri'
    $range2.Font.Bold = $true
    $range2.Font.ColorIndex = '-4105'
    $range3 = $xl.Range("K1").Entirecolumn
    $range3.Select()
    $xlfilter = "Arrived"
    $range3.AutoFilter(11, $xlfilter)
    $rangefinal = $xl.Range("A1")
    $rangefinal.Select()

    $wb.save()
    $xl.Workbooks.Close()
    $xl.Quit()

    Get-process EXCEL | stop-process
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
   
  move-item -Path $a `
               -Destination '\\wserver\Kiosk\Invoices\Hylton Ross'
    }
 else {
    Guy-SendGmail "Copy of Hylton Ross commission spreadsheet failed" "PLEASE INVESTIGATE"
 }