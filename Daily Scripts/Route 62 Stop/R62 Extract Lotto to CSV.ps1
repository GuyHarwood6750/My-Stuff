<#      Extract from Customer spreadsheet the range for new invoices to be generated.
        Modify the $StartR (startrow) and $endR (endrow).
        This can only be done by eyeball as spreadsheet has historical data.
#>

$inspreadsheet = 'C:\userdata\route 62\_Development\airtime - Electricty - Lotto monthly breakdown.xlsx'        #Change each month
$csvfile = 'lotto.csv'                            #Temp file
$pathout = 'C:\userdata\route 62\_Development\'
$custsheet = 'May 2020'                          #Customer worksheet - change each month
$outfile2 = 'C:\userdata\Route 62\_Development\lotto2.csv' #Change each month
$startR = 47                                    #Start row       
$endR = 50                                    #End Row
$startCol = 10                                    #Start Col
$endCol = 18                                      #End Col

$Outfile = $pathout + $csvfile

Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol -NoHeader| Select-Object * | Export-Csv -Path $Outfile -NoTypeInformation

# Format date column correctly
        Get-ChildItem -Path $pathout -Name $csvfile
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $xl.workbooks.Open($Outfile)
        $xl.Sheets.Item('lotto').Activate()
  
        $range = $xl.Range("a:a").Entirecolumn
        $range.NumberFormat = 'dd/mm/yyyy'

        $wb.save()
        $xl.Workbooks.Close()
        $xl.Quit()

Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
#Remove-Item -Path $outfile