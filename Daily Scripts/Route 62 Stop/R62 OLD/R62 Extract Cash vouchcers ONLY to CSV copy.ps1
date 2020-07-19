<#      Extract from cash vouchers from spreadsheet the range for new invoices to be generated.
        Modify the $StartR (startrow) and $endR (endrow).
        This can only be done by eyeball as spreadsheet has historical data.
#>

$inspreadsheet = 'C:\userdata\route 62\_All Suppliers\Suppliers June 2020.xlsm'
$csvfile = 'SHEET1.csv'
$pathout = 'C:\userdata\route 62\_All Suppliers\'
$custsheet = 'June 2020'                                        #Month worksheet
$outfile2 = 'C:\userdata\route 62\_All Suppliers\CSH June 2020.csv'
$startR = 5                                             #Start row - do not change
$endR = 22                                              #End Row - changes each month depending on number of purchases
$startCol = 1                                           #Start Col (don't change)
$endCol = 9                                             #End Col (don't change)
$filter = "CSH"

$Outfile = $pathout + $csvfile

Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol -NoHeader| Where-Object -Filterscript {$_.P2 -eq $filter} | Export-Csv -Path $Outfile -NoTypeInformation

        Get-ChildItem -Path $pathout -Name $csvfile
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $xl.workbooks.Open($Outfile)
        $xl.Sheets.Item('sheet1').Activate()
        $range = $xl.Range("d:d").Entirecolumn
        $range.NumberFormat = 'dd/mm/yyyy'

        $wb.save()
        $xl.Workbooks.Close()
        $xl.Quit()

Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
Remove-Item -Path $outfile
#>