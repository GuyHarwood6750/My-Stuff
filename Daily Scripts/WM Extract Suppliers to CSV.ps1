<#      Extract from Customer spreadsheet the range for new invoices to be generated.
        Modify the $StartR (startrow) and $endR (endrow).
        This can only be done by eyeball as spreadsheet has historical data.
#>

$inspreadsheet = 'C:\userdata\circe launches\_all suppliers\supplier invoices cash vouchers 2021.xlsx'        #Change each month
$csvfile = 'suppliers_1.csv'                            #Temp file
$pathout = 'C:\userdata\circe launches\_all suppliers\'
$custsheet = 'May 2020'                          #Customer worksheet - change each month
$outfile2 = 'C:\userdata\circe launches\_all suppliers\supplier may 2020.csv' #Change each month
$startR = 5                                    #Start row - does not change       
$endR = 9                                    #End Row - changes each month depending on number of invoices
$startCol = 1                                    #Start Col (don't change)
$endCol = 8                                      #End Col (don't change)

$Outfile = $pathout + $csvfile

Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol -NoHeader| Select-Object P1, P2, P3, P4, P5, P6, P7, P8 | Export-Csv -Path $Outfile -NoTypeInformation

# Format date column correctly
        Get-ChildItem -Path $pathout -Name $csvfile
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $xl.workbooks.Open($Outfile)
        $xl.Sheets.Item('suppliers_1').Activate()
  
        $range = $xl.Range("c:c").Entirecolumn
        $range.NumberFormat = 'dd/mm/yyyy'

        $wb.save()
        $xl.Workbooks.Close()
        $xl.Quit()

Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
Remove-Item -Path $outfile