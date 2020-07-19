<#  Warren Marine
    Get list of Cash Vouchers
    Output to text file to be imported as a Pastel Cashbook batch.
#>
#Input from Supplier spreadsheet
$csvclient = 'C:\Userdata\Circe Launches\_ALL SUPPLIERS\CSH June 2020.csv'                  
#Temp file
$outfile1 = 'C:\Userdata\Circe Launches\_ALL SUPPLIERS\cashvch1.txt'                  
#File to be imported into Pastel
$outfileF = 'C:\Userdata\Circe Launches\_ALL SUPPLIERS\cashvchpastel.txt'             
#Remove last file imported to Pastel
$checkfile = Test-Path $outfileF
if ($checkfile) { Remove-Item $outfilef }                   

#Import latest csv from Supplier spreadsheet
$data = Import-Csv -path $csvclient -header acc, Expacc, date, ref, invnum, desc, amt, vat    

foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods -transactiondate $aObj.date

    Switch ($aObj.Expacc) {
        CLN {$expacc = '3210000'}            #Cleaning
        FUEL {$expacc = '4150000' }         #Motor vehicles
        GIFT {$expacc = '3551000'}            #Trade Gifts
        MED {$expacc = '4500000'}            #Medical expenses, Staff welfare
        MVE {$expacc = '4150000' }         #Motor vehicles
        PC {$expacc = '4550000'}            #Protective clothing
        RENT {$expacc = '4300000'}            #Rent
        RM {$expacc = '4350000'}            #Repairs and Maintenance
        REF {$expacc = '4500000'}            #Staff refreshments
        SS {$expacc = '3750000'}            #Ship stores & provisions
        STATIONARY {$expacc = '4200000'}    #Stationery
        TEL {$expacc = '4600000'}            #Telephone
        TETA {$expacc = '4451000'}            #TETA Training
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