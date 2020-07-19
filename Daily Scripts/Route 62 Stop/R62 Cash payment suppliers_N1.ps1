<#      Extract from Customer spreadsheet the range for new invoices to be generated.
        Modify the $StartR (startrow) and $endR (endrow).
        This can only be done by eyeball as spreadsheet has historical data.
#>

$inspreadsheet = 'C:\userdata\route 62\_all suppliers\suppliers July 2020.xlsm'
$csvfile = 'suppliers cash payment.csv'
$pathout = 'C:\userdata\route 62\_all suppliers\'
$custsheet = 'july 2020'                          #Customer worksheet
$outfile2 = 'C:\userdata\Route 62\_aLL suppliers\suppliers paid cash july 2020.csv'
$startR = 5                                    #Start row
$endR = 80                                    #End Row
$startCol = 1                                    #Start Col (don't change)
$endCol = 11                                      #End Col (don't change)
$filter = "CSH"

$Outfile = $pathout + $csvfile

Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol -NoHeader -DataOnly | Where-Object -FilterScript {$_.P2 -ne $filter -and $_.P2 -ne 'CC' -and $_.P10 -eq 'cash' -and $_.P11 -ne 'done'} | Export-Csv -Path $Outfile -NoTypeInformation

# Format date column correctly
Get-ChildItem -Path $pathout -Name $csvfile
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.workbooks.Open($Outfile)
$xl.Sheets.Item('suppliers cash payment').Activate()
  
$range = $xl.Range("d:d").Entirecolumn
$range.NumberFormat = 'dd/mm/yyyy'

$wb.save()
$xl.Workbooks.Close()
$xl.Quit()

Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
Remove-Item -Path $outfile
<#  Get list of Cash payments to suppliers
    Output to text file to be imported as a Pastel Cashbook batch.
#>
 #Input from Supplier spreadsheet
#$csvclient = 'C:\Userdata\route 62\_all suppliers\suppliers paid cash.csv'                 
$outfile1 = 'C:\Userdata\route 62\_all suppliers\cashpur1.txt'                  #Temp file
#File to be imported into Pastel
$outfileF = 'C:\Userdata\route 62\_all suppliers\cashsuppliers.txt'             
#Remove last file imported to Pastel
$checkfile = Test-Path $outfileF
if ($checkfile) { Remove-Item $outfilef }                   

#Import latest csv from Supplier spreadsheet
$data = Import-Csv -path $outfile2 -header alloc, suppacc, desc, date, ref, invnum, descr2, amt    

foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods -transactiondate $aObj.date
    
       #Format Pastel batch
    
    $props1 = [ordered] @{
        Period  = $pastelper
        Date    = $aObj.date
        GL      = 'C'
        contra  = $aobj.suppacc                       #Expense account to be debited (DR)
        ref     = $aObj.ref
        comment = $aObj.desc
        amount  = $aObj.amt
        fil1    = '0'
        fil2    = '0'
        fil3    = ' '
        fil4    = ' '
        fil5    = '8430000'                     #Cash voucher contra account number
        fil6    = '1'
        fil7    = '1'
        fil8    = '0'
        fil9    = '0'
        fil10   = '0'
        amt2    = $aObj.amt
    }
      
        $objlist = New-Object -TypeName psobject -Property $props1
        $objlist | Select-Object * | Export-Csv -path $outfile1 -NoTypeInformation -Append
    
    }  
    #Remove header information so file can be imported into Pastel Accounting.
    
    Get-Content -Path $outfile1 | Select-Object -skip 1 | Set-Content -path $outfilef
    Remove-Item -Path $outfile1