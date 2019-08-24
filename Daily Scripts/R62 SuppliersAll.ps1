<#  Get list of Supplier invoices from spreadsheet
    Output to text file to be imported as a Pastel Invoice batch.

#>
$csvclient = 'C:\userdata\route 62\_all suppliers\suppliers 16AUG.csv'      #Input from Client spreadsheet
$outfile = 'C:\userdata\route 62\_all suppliers\supplierinv.txt'        #Temp file
$outfile2 = 'C:\userdata\route 62\_all suppliers\suppliersAUG16.txt'     #File to be imported into Pastel

#Remove last file imported to Pastel

$checkfile = Test-Path $outfile2
if ($checkfile) { Remove-Item $outfile2 }                   

#Import latest csv from Client spreadsheet

$data = Import-Csv -path $csvclient -header acc, date, invnum, amt, vat, method

foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods -transactiondate $aObj.date

    [decimal]$amount = $aObj.amt
    [decimal]$vat = $amount * 15 / 115
    [decimal]$amtexvat = $aObj.amt - $vat
    $vatexamt = [math]::Round($amtexvat, 2)

    Switch ($aObj.acc) {
        BADEN1 { $expacc = '2000010'; $description = 'Purchases' }
        BAT { $expacc = '2000010'; $description = 'Purchases' }
        CAPDRY { $expacc = '2000010'; $description = 'Purchases' }
        GAYD { $expacc = '2000010'; $description = 'Purchases' }
        GEI { $expacc = '2000010'; $description = 'Purchases' }
        HENT { $expacc = '2000010'; $description = 'Purchases' }
        JANIBA { $expacc = '2000010'; $description = 'Purchases' }
        KKI { $expacc = '2000010'; $description = 'Purchases' }
        KONIG { $expacc = '2000010'; $description = 'Purchases' }
        KOW { $expacc = '2000010'; $description = 'Purchases' }
        MANNY { $expacc = '2000010'; $description = 'Purchases' }
        MBRAND { $expacc = '2000010'; $description = 'Purchases' }
        MERRYP { $expacc = '2000010'; $description = 'Purchases' }
        NUTS { $expacc = '2000010'; $description = 'Purchases' }
        OFOOD { $expacc = '2000010'; $description = 'Purchases' }
        PACKT { $expacc = '2000010'; $description = 'Purchases' }
        PARMA { $expacc = '2000010'; $description = 'Purchases' }
        PENBEV { $expacc = '2000010'; $description = 'Purchases' }
        RAIM { $expacc = '2000010'; $description = 'Purchases' }
        SGV { $expacc = '2000010'; $description = 'Purchases' }
        SIMBA { $expacc = '2000010'; $description = 'Purchases' }
        TWISP { $expacc = '2000010'; $description = 'Purchases' }
        VASC { $expacc = '2000010'; $description = 'Purchases' }
        VEGA { $expacc = '2000010'; $description = 'Purchases' }
        
        Default { $expacc = '2000010' }
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
        f16   = '30/04/2019'
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
        f31   = '15'
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
Remove-Item -Path $outfile
#Remove-Item -Path $csvclient