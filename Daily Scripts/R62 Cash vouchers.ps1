<#  Get list of Cash purchases
    Output to text file to be imported as a Pastel Cashbook batch.
#>
 #Input from Client spreadsheet
$csvclient = 'C:\Userdata\route 62\cash purchases\cash purchases september 2019.csv'                 
$outfile1 = 'C:\Userdata\route 62\cash purchases\cashsupplier.txt'                  #Temp file
#File to be imported into Pastel
$outfileF = 'C:\Userdata\route 62\cash purchases\cashpurpastel.txt'             
#Remove last file imported to Pastel
$checkfile = Test-Path $outfileF
if ($checkfile) { Remove-Item $outfilef }                   

#Import latest csv from Client spreadsheet
$data = Import-Csv -path $csvclient -header Expacc, date, ref, desc, amt, vat    

foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods -transactiondate $aObj.date
    
    Switch ($aObj.Expacc) {
        ADV { $expacc = '3050000' }         
        CLN { $expacc = '3250000' }         
        COMPA { $expacc = '6250010' }         
        COUR { $expacc = '3400000' }
        ELEC { $expacc = '3650000' }
        FUEL { $expacc = '4150001' }
        MVR { $expacc = '4150002' }
        NPUR { $expacc = '2000012' }
        PUR {$expacc = '2000010'}         
        PVT { $expacc = '5201001' }         
        RM { $expacc = '4350000' } 
        STAT { $expacc = '4200000' }
        TEL { $expacc = '4600000' }         
        
        Default {$expacc = '9983000'}       
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