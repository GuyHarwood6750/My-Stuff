<#  Get list of Cash payments to suppliers
    Output to text file to be imported as a Pastel Cashbook batch.
#>
 #Input from Client spreadsheet
$csvclient = 'C:\Userdata\route 62\cash supplier payments\suppliers february 2020.csv'                 
$outfile1 = 'C:\Userdata\route 62\cash supplier payments\cashpur1.txt'                  #Temp file
#File to be imported into Pastel
$outfileF = 'C:\Userdata\route 62\cash supplier payments\cashsuppliers.txt'             
#Remove last file imported to Pastel
$checkfile = Test-Path $outfileF
if ($checkfile) { Remove-Item $outfilef }                   

#Import latest csv from Client spreadsheet
$data = Import-Csv -path $csvclient -header suppacc, date, ref, desc, amt    

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