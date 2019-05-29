<#  Get list of invoices from spreadsheet
    Output to text file to be imported as a Pastel Invoice batch.

#>
$csvclient = 'C:\test\invWM\invsam22may.csv'            #Input csv file
$csvrate = 'C:\test\invwm\accitemrate.csv'      #Rates per customer
$outfile = 'C:\test\invWM\WMinvTmp.txt'        #Temp file
$outfile2 = 'C:\test\invWM\WMinvioces.txt'     #File to be imported into Pastel
#Remove last file imported to Pastel

$checkfile = Test-Path $outfile2
if ($checkfile) { Remove-Item $outfile2 }                   

#Import latest csv from Client spreadsheet

$data = Import-Csv -path $csvclient -header accnum, date, time, customername, BookingID, groupname, voucher, ptype, type, qty, rate, guide


$ratelkup = Import-csv -Path $csvrate -Header acc, code, desc, rate

foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods -transactiondate $aObj.date
    
    foreach ($bObj in $ratelkup) {
        
        if ($aObj.'accnum' -eq $bObj.'acc') {
            $ratecode = $bObj.code
            $description = $bObj.desc
            [decimal]$amount = $bObj.rate
            [decimal]$vat = $amount * 15 / 115
            [decimal]$amtexvat = $bObj.rate - $vat
            $vatexamt = [math]::Round($amtexvat, 2)
         }
        else {
        }
    }
    #Format Pastel batch
    $props = [ordered] @{
        hd    = 'Header'
        f1    = ''
        f2    = ''
        f3    = 'Y'
        acc   = $aObj.accnum
        per   = $pastelper
        dte   = $aObj.date
        order = $aObj.groupname
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
        f16   = '30/06/2019'
        f17   = ''
        f18   = ''
        f19   = ''
        f20   = '1'
        f21   = ''
        f22   = ''
        f23   = ''
        f24   = 'Y'
        f25   = 'Detail'
        f26   = '0'
        f27   = $aObj.qty
        f28   = $vatexamt
        f29   = $amount
        f30   = ''
        f31   = '15'
        f32   = '3'
        f33   = '0'
        f34   = $ratecode                    #rate code
        f35   = $description
        f36   = '4'
        f37   = ''
        f38   = ''
        f39   = 001
        #
        f53   = 'Detail'
        f54   = '0'
        f55   = '1'
        f56   = '0'
        f57   = '0'
        f58   = ''
        f59   = '0'
        f60   = '3'
        f61   = '0'
        f62   = "'"
        f63   = $aObj.groupname
        f64   = 7
        f65   = ''
        f66   = ''
        #
        f67   = 'Detail'
        f68   = '0'
        f69   = '1'
        f70   = '0'
        f71   = '0'
        f72   = ''
        f73   = '0'
        f74   = '3'
        f75   = '0'
        f76   = "'"
        f77   = $aObj.voucher
        f78   = 7
        f79   = ''
        f80   = ''
        #
        
        f67a  = 'Detail'
        f68b  = '0'
        f69c  = '1'
        f70d  = '0'
        f71e  = '0'
        f72f  = ''
        f73g  = '0'
        f74h  = '3'
        f75i  = '0'
        f76j  = "'"
        f77k  = 'Date: ' + ' ' + $aObj.date
        f78l  = 7
        f79m  = ''
        f80n  = ''
        #
        f81   = 'Detail'
        f82   = '0'
        f83   = '1'
        f84   = '0'
        f85   = '0'
        f86   = ''
        f87   = '0'
        f88   = '3'
        f89   = '0'
        f90   = "'"
        f91   = 'Time: ' + ' ' + $aObj.time
        f92   = 7
        f93   = ''
        f94   = ''
        #
        f95   = 'Detail'
        f96   = '0'
        f97   = '1'
        f98   = '0'
        f99   = '0'
        f100  = ''
        f101  = '0'
        f102  = '3'
        f103  = '0'
        f104  = "'"
        f105  = 'Guide: ' + ' ' + $aObj.guide
        f106  = 7
        f107  = ''
        f108  = ''
        #
        f95a   = 'Detail'
        f96b   = '0'
        f97c   = '1'
        f98d   = '0'
        f99e   = '0'
        f100f  = ''
        f101g  = '0'
        f102h  = '3'
        f103i  = '0'
        f104j  = "'"
        f105k  = 'Our ref: ' + ' ' + $aObj.BookingID
        f106l  = 7
        f107m  = ''
        f108n  = ''

    }
                 
    $objlist = New-Object -TypeName psobject -Property $props 
    $objlist | Select-Object * | Export-Csv -path $outfile -NoTypeInformation -Append
}  
#Remove header information so file can be imported into Pastel Accounting.

Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
Remove-Item -Path $outfile
#Remove-Item -Path $csvclient