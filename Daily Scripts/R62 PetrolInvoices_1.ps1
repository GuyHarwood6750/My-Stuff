<#      Extract from Customer spreadsheet the range for new invoices to be generated.
        Modify the $StartR (startrow) and $endR (endrow).
        This can only be done by eyeball as spreadsheet has historical data.
#>

$inspreadsheet = 'C:\userdata\route 62\petrol books\PETROL BOOKS.xlsx'
$csvfile = 'sheet1.csv'
$pathout = 'C:\userdata\route 62\petrol books\'
$custsheet = 'WWB'                          #Customer worksheet
$startR = 3161                                    #Start row
$endR = 3166                                     #End Row
$startCol = 1                                    #Start Col (don't change)
$endCol = 9                                      #End Col (don't change)

$Outfile = $pathout + $csvfile

Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol | Export-Csv -Path $Outfile -NoTypeInformation

# Format date column correctly
        Get-ChildItem -Path $pathout -Name $csvfile
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $xl.workbooks.Open($Outfile)
        $xl.Sheets.Item('sheet1').Activate()
  
        $range = $xl.Range("b:b").Entirecolumn
        $range.NumberFormat = 'dd/mm/yyyy'

        $wb.save()
        $xl.Workbooks.Close()
        $xl.Quit()