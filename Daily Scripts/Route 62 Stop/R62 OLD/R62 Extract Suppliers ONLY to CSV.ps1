<#      Extract from EXPENSES spreadsheet the range for new invoices to be generated.
        Modify the $StartR (startrow) and $endR (endrow). 
#>
$inspreadsheet = 'C:\userdata\route 62\_all suppliers\suppliers june 2020.xlsm'          #Source workbook
$csvfile = 'suppliers_1.csv'                                                                                    #Temp file
$pathout = 'C:\userdata\route 62\_all suppliers\'
$custsheet = 'june 2020'                                                                        #Month worksheet - changes each month
$outfile2 = 'C:\userdata\route 62\_all suppliers\suppliers june 2020.csv'                  #Change each month
$startR = 5                                             #Start row - does not change       
$endR = 22                                              #End Row - changes each month depending on number of invoices
$startCol = 1                                           #Start Col (don't change)
$endCol = 9                                             #End Col (don't change)
$filter= "CSH"                                          #Filter - Not CASH VOUCHERS - SER Where-Object BELOW
$Outfile = $pathout + $csvfile

Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol -NoHeader | Where-Object -Filterscript { $_.P2 -ne $filter} | Export-Csv -Path $Outfile -NoTypeInformation

# Format date column correctly
        Get-ChildItem -Path $pathout -Name $csvfile
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $xl.workbooks.Open($Outfile)
        $xl.Sheets.Item('suppliers_1').Activate()
        $range = $xl.Range("d:d").Entirecolumn
        $range.NumberFormat = 'dd/mm/yyyy'

        $wb.save()
        $xl.Workbooks.Close()
        $xl.Quit()

Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
Remove-Item -Path $outfile