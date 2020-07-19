<#  GUY\_SARS
    Get list of bank transactions (csv) captured manually from paper.
    Output to text file to be imported as a Pastel Cashbook batch.
    Output file for payments (P)
    Output file for receipts (R)
#>
#Input bank transactions captured manually
$csvclient = 'C:\Userdata\GUY\_SARS\2019\statements\bank transactions\d0212t.csv'                  
#Temp file
$outfile1 = 'C:\Userdata\GUY\_SARS\2019\statements\bank transactions\tempreceipts.txt'                  
$outfile2 = 'C:\Userdata\GUY\_SARS\2019\statements\bank transactions\temppayments.txt'                  
#File to be imported into Pastel
$outfileF1 = 'C:\Userdata\GUY\_SARS\2019\statements\bank transactions\d0212tR.txt'       #Receipts           
$outfileF2 = 'C:\Userdata\GUY\_SARS\2019\statements\bank transactions\d0212tP.txt'       #Payments            
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
    $pastelper = PastelPeriods -transactiondate $aObj.date
    Switch ($aObj.acc) {
        AGRAY { $acc = '8300000' }
        AGRAYA { $acc = '1000006' }
        BC { $acc = '3200000' }
        BKSOFTL { $acc = '4550000' }
        COMP { $acc = '3300000' }
        DFOOD { $acc = '3600001' }
        DISC { $acc = '8420000' }
        DIV { $acc = '1000009' }
        DON { $acc = '3550000' }
        ELEC { $acc = '3650000' }
        FNBC { $acc = '8410000' }
        FT { $acc = '4450000' }
        INSREF { $acc = '2850000' }
        INTP { $acc = '3900000' }
        INTR { $acc = '2750000' }
        KUSA { $acc = '4450000' }
        LOT { $acc = '4500000' }
        MEDP { $acc = '4000000' }
        MVR { $acc = '4150002' }
        PHSOFTL { $acc = '4551000' }
        POL { $acc = '3950000' }
        PTAX { $acc = '3800004' }
        PVT { $acc = '4500000' }
        RA { $acc = '3951000' }
        REP { $acc = '4350000' }
        SA { $acc = '4050000' }
        SARSR { $acc = '3800005' }
        SEC { $acc = '4210000' }
        STA { $acc = '4201000' }
        SWM { $acc = '3650000' }
        TI { $acc = '4600000' }
        UINC { $acc = '9992000' }
        UINP { $acc = '9991000' }
        VET { $acc = '3600002' }
        W { $acc = '4400000' }
        WSAR { $acc = '4451000' }
        Default { $acc = '9993000' }
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
                fil5    = '8420000'                #Bank account number
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
                fil5    = '8420000'                     #Bank account number
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