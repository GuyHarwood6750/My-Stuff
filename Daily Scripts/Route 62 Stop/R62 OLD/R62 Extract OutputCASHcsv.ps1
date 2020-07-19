<#      Extract from Customer spreadsheet the range for new invoices to be generated.
        Modify the $StartR (startrow) and $endR (endrow).
        This can only be done by eyeball as spreadsheet has historical data.
#>

$inspreadsheet = 'C:\userdata\route 62\_all suppliers\suppliers May 2020.xlsm'
$csvfile = 'suppliers cash payment.csv'
$pathout = 'C:\userdata\route 62\_all suppliers\'
$custsheet = 'outputcashcsv'                          #Customer worksheet
$outfile2 = 'C:\userdata\Route 62\_aLL suppliers\suppliers paid cash.csv'
$startR = 2                                    #Start row
$endR = 22                                    #End Row
$startCol = 1                                    #Start Col (don't change)
$endCol = 9                                      #End Col (don't change)

$Outfile = $pathout + $csvfile

Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol -NoHeader| Select-Object P1, P2, P3, P4, P5, P6, P7 | Export-Csv -Path $Outfile -NoTypeInformation

# Format date column correctly
        Get-ChildItem -Path $pathout -Name $csvfile
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $xl.workbooks.Open($Outfile)
        $xl.Sheets.Item('suppliers cash payment').Activate()
  
        $range = $xl.Range("c:c").Entirecolumn
        $range.NumberFormat = 'dd/mm/yyyy'

        $wb.save()
        $xl.Workbooks.Close()
        $xl.Quit()

Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
Remove-Item -Path $outfile