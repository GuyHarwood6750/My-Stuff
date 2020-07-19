<#      Extract from cash vouchers from spreadsheet the range for new invoices to be generated.
        Modify the $StartR (startrow) and $endR (endrow).
        This can only be done by eyeball as spreadsheet has historical data.
#>

$inspreadsheet = 'C:\userdata\circe launches\_All Suppliers\Supplier invoices cash vouchers 2021.xlsx'
$csvfile = 'SHEET1.csv'
$pathout = 'C:\userdata\circe launches\_All Suppliers\'
$custsheet = 'June 2020'                                        #Month worksheet
$outfile2 = 'C:\userdata\circe launches\_All Suppliers\CSH June 2020.csv'
$startR = 2                                             #Start row - do not change
$endR = 34                                              #End Row - changes each month depending on number of purchases
$startCol = 1                                           #Start Col (don't change)
$endCol = 10                                             #End Col (don't change)
$filter = "CSH"

$Outfile = $pathout + $csvfile

Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol -NoHeader -DataOnly| Where-Object -Filterscript {$_.P1 -eq $filter -and $_.P10 -ne 'done'} | Export-Csv -Path $Outfile -NoTypeInformation

        Get-ChildItem -Path $pathout -Name $csvfile
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $xl.workbooks.Open($Outfile)
        $xl.Sheets.Item('sheet1').Activate()
        $range = $xl.Range("c:c").Entirecolumn
        $range.NumberFormat = 'dd/mm/yyyy'

        $wb.save()
        $xl.Workbooks.Close()
        $xl.Quit()

Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
Remove-Item -Path $outfile
#>