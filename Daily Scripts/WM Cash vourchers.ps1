<#  Warren Marine
    Get list of Cash Vouchers
    Output to text file to be imported as a Pastel Cashbook batch.
#>
#Input from Client spreadsheet
$csvclient = 'C:\Userdata\Circe Launches\cash vouchers\cv24july.csv'                  
#Temp file
$outfile1 = 'C:\Userdata\Circe Launches\cash vouchers\cashvch1.txt'                  
#File to be imported into Pastel
$outfileF = 'C:\Userdata\Circe Launches\cash vouchers\cashvchpastel.txt'             
#Remove last file imported to Pastel
$checkfile = Test-Path $outfileF
if ($checkfile) { Remove-Item $outfilef }                   

#Import latest csv from Client spreadsheet
$data = Import-Csv -path $csvclient -header Expacc, date, ref, desc, amt, vat    

foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods -transactiondate $aObj.date

    Switch ($aObj.Expacc) {
        MVE {$expacc = '4150000' }         #Motor vehicles
        RM {$expacc = '4350000'}            #Repairs and Maintenance
        TEL {$expacc = '4600000'}            #Telephone
        CLN {$expacc = '3210000'}            #Cleaning
        PC {$expacc = '4550000'}            #Protective clothing
        MED {$expacc = '4050000'}            #Medical expenses
        SS {$expacc = '3750000'}            #Ship stores & provisions
        Default {$expacc = '9992000'}       #Unallocated Expense account
    }
    Switch ($aObj.vat) {
        Y { $VATind = '15' }
        N { $VATind = '0' }
        Default {$VATind = '15'}
    }
    #Format Pastel batch
    $props1 = [ordered] @{
        Period  = $pastelper
        Date    = $aObj.date
        GL      = 'G'
        contra  = $expacc                       #Expense account to be debited (DR)
        ref     = $aObj.ref
        comment = $aObj.desc
        amount  = $aObj.amt
        fil1    = $VATind
        fil2    = '0'
        fil3    = ' '
        fil4    = '     '
        fil5    = '8410000'                     #Cash voucher contra account number
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