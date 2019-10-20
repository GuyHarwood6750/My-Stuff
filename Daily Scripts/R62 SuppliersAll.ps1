﻿<#  Get list of Supplier invoices from spreadsheet
    Output to text file to be imported as a Pastel Invoice batch.
#>
#Input from Supplier spreadsheet
$csvsupplier = 'C:\userdata\route 62\_all suppliers\september_d.csv'
#Temp file      
$outfile = 'C:\userdata\route 62\_all suppliers\supplierinv.txt'
#File to be imported into Pastel        
$outfile2 = 'C:\userdata\route 62\_all suppliers\september_d.txt'     

#Remove last file imported to Pastel
$checkfile = Test-Path $outfile2
if ($checkfile) { Remove-Item $outfile2 }                   

#Import latest csv from Supplier spreadsheet, VAT & NO-VAT, not MIXED VAT.
$data = Import-Csv -path $csvsupplier -header acc, date, invnum, descr, amt, vat

foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods -transactiondate $aObj.date

    #Process transactions based on VAT=Y or VAT=N  
    switch ($aObj.vat) { 
        Y {
            [decimal]$amount = $aObj.amt
            [decimal]$vat = $amount * 15 / 115
            [decimal]$amtexvat = $aObj.amt - $vat
            $vatexamt = [math]::Round($amtexvat, 2)
            $vatpercent = 15 
            $expacc = '2000010'
            $description = $aObj.descr
        }
        N {
            [decimal] $amount = $aObj.amt
            [decimal] $vatexamt = $aObj.amt
            $vatpercent = 0 
            $expacc = '2000012' 
            $description = $aObj.descr
        }
    }   
    #Process Supplier that are not 'default purchases'
    Switch ($aObj.acc) {
        AIDOR { $expacc = '4350000'; $description = $aObj.descr }
        AUTOC { $expacc = '4150002'; $description = $aObj.descr }
        CON001 { $expacc = '4600000'; $description = $aObj.descr }
        GRIDH { $expacc = '4600000'; $description = $aObj.descr }
        METRA { $expacc = '4350000'; $description = $aObj.descr }
        MOOVR { $expacc = '4300000'; $description = $aObj.descr }
        MIOSA { $expacc = '4550000'; $description = $aObj.descr }
        MSCHER { $expacc = '3000000'; $description = $aObj.descr }
        RENOKI { $expacc = '3250000'; $description = $aObj.descr }
        SAMRO { $expacc = '4550000'; $description = $aObj.descr }
        SWDMUN { $expacc = '3650000'; $description = $aObj.descr }
        WAF00 { $expacc = '4600000'; $description = $aObj.descr }
        WALTON { $expacc = '4200000'; $description = $aObj.descr }
        
    }
    
    #Format Pastel batch
    $props = [ordered] @{
        hd    = 'Header'
        f1    = ''
        f2    = ''
        f3    = ''
        acc   = $aObj.acc
        per   = $pastelper
        dte   = $aObj.date
        order = $aObj.invnum
        f4    = "Y"
        f5    = '0'
        f6    = ''
        f7    = ''
        f8    = ''
        f9    = ''
        f10   = ''
        f11   = ''
        f12   = ''
        f13   = ''
        f14   = ''
        f15   = '0'
        f16   = $aObj.date
        f17   = ''
        f18   = ''
        f19   = ''
        f20   = '1'
        f21   = ''
        f22   = ''
        f23   = 'N'
        f24   = ''
        f25   = 'Detail'
        f26   = $vatexamt
        f27   = '1'
        f28   = $vatexamt
        f29   = $aObj.amt
        f30   = ''
        f31   = $vatpercent
        f32   = '0'
        f33   = '0'
        f34   = $expacc
        f35   = $description
        f36   = '6'
        f37   = ''
        f38   = ''
    }
                 
    $objlist = New-Object -TypeName psobject -Property $props 
    $objlist | Select-Object * | Export-Csv -path $outfile -NoTypeInformation -Append
}  
#Remove header information so file can be imported into Pastel Accounting.
Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
#Remove Temp file.
Remove-Item -Path $outfile