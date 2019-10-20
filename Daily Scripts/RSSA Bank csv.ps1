<#  RSSA
    Get list of bank transactions (csv) captured manually from paper.
    Output to text file to be imported as a Pastel Cashbook batch.
    Output file for payments (P)
    Output file for receipts (R)
#>
#Input bank transactions captured manually
$csvclient = 'C:\Userdata\RSSA\2019\bank transactions\899.csv'                  
#Temp file
$outfile1 = 'C:\Userdata\RSSA\2019\bank transactions\tempreceipts.txt'                  
$outfile2 = 'C:\Userdata\RSSA\2019\bank transactions\temppayments.txt'                  
#File to be imported into Pastel
$outfileF1 = 'C:\Userdata\RSSA\2019\bank transactions\899R.txt'       #Receipts           
$outfileF2 = 'C:\Userdata\RSSA\2019\bank transactions\899P.txt'       #Payments            
#Remove last file imported to Pastel
$checkfile = Test-Path $outfileF1
if ($checkfile) { Remove-Item $outfilef1 }                   
$checkfile = Test-Path $outfileF2
if ($checkfile) { Remove-Item $outfilef2 }                   
#
#Import latest csv from Client spreadsheet
$data = Import-Csv -path $csvclient -header type, acc, date, ref, desc, amt    
#
foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods2 -transactiondate $aObj.date
    Switch ($aObj.acc) {
        ACCOM { $acc = '3950030' }
        ACC { $acc = '3000000' }
        ASSAF { $acc = '1600010' }
        AGRAY { $acc = '7100000' }
        AWARDS { $acc = '4210010' }
        BC { $acc = '3200000' }
        CASH { $acc = '9984000' }
        CLAUDE { $acc = '1300040' }
        CLR { $acc = '9984000' }
        COMPA { $acc = '6250010' }
        DINE { $acc = '3910000' }
        DINI { $acc = '1400000' }
        DIV { $acc = '2760000' }
        MARK { $acc = '4620020' }
        ML { $acc = '8410000' }
        PACFEE { $acc = '9400020' }
        PAYE { $acc = '4400030' }
        PRINT { $acc = '4200000' }
        PRIZE { $acc = '4620010' }
        PSAL { $acc = '9400010' }
        REF { $acc = '3950010' }
        SAL { $acc = '4400010' }
        SP { $acc = '4352000' }
        SUBS { $acc = '1000000' }
        TRANS { $acc = '4610020' }
        TRAVEL { $acc = '3950020' }
        UNKR { $acc = '9991000' }
        UNKP { $acc = '9992000' }
        WEB { $acc = '4601000' }
        Default { $acc = '9983000' }
    }
    Switch ($aObj.type) {
        r {  
            #Format Pastel receipt batch
            $props1 = [ordered] @{
                Period  = $pastelper
                Date    = $aObj.date
                GL      = 'G'
                contra  = $acc                    #account to be debited (DR) or credited (CR)
                ref     = $aObj.ref
                comment = $aObj.desc
                amount  = $aObj.amt
                fil1    = '0'
                fil2    = '0'
                fil3    = ' '
                fil4    = '     '
                fil5    = '8400000'                #Bank account number
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
        p {
            #Format Pastel payments batch
            $props2 = [ordered] @{
                Period  = $pastelper
                Date    = $aObj.date
                GL      = 'G'
                contra  = $acc                    #account to be debited (DR) or credited (CR)
                ref     = $aObj.ref
                comment = $aObj.desc
                amount  = $aObj.amt
                fil1    = '0'
                fil2    = '0'
                fil3    = ' '
                fil4    = '     '
                fil5    = '8400000'                     #Bank account number
                fil6    = '1'
                fil7    = '1'
                fil8    = '0'
                fil9    = '0'
                fil10   = '0'
                amt2    = $aObj.amt
            }
            $objlist = New-Object -TypeName psobject -Property $props2
            $objlist | Select-Object * | Export-Csv -path $outfile2 -NoTypeInformation -Append
        }
    }
    }  
    #Remove header information so file can be imported into Pastel Accounting.
    Get-Content -Path $outfile1 | Select-Object -skip 1 | Set-Content -path $outfilef1
    Get-Content -Path $outfile2 | Select-Object -skip 1 | Set-Content -path $outfilef2
    Remove-Item -Path $outfile1
    Remove-Item -Path $outfile2