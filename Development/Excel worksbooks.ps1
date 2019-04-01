$path = 'C:\test\'
$file = Get-ChildItem -Path $path -Name 'HrossBookingDetails*.xlsx'
$a = $path + $file

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.workbooks.Open($a)
$ws = $xl.Sheets.Item('Sheet1').Activate()
#$ws = $wb.Sheets.Add()
#$ws1 = $wb.Sheets.Item("sheet2").delete()
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
$range3.AutoFilter(11,$xlfilter)

$wb.save()
$xl.Workbooks.Close()
$xl.Quit()

Get-process EXCEL | stop-process
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)


