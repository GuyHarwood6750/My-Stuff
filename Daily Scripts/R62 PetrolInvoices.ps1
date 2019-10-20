<#  Get list of petrol invoices from Petrol books spreadsheet
    Output to text file to be imported as a Pastel Invoice batch.

#>
$csvclient = 'C:\userdata\route 62\petrol books\all17oct.csv'      #Input from Client spreadsheet
$outfile = 'C:\userdata\route 62\petrol books\petrolinv.txt'        #Temp file
$outfile2 = 'C:\userdata\route 62\petrol books\all17oct.txt'     #File to be imported into Pastel

#Remove last file imported to Pastel

$checkfile = Test-Path $outfile2
if ($checkfile) { Remove-Item $outfile2 }                   

#Import latest csv from Client spreadsheet

$data = Import-Csv -path $csvclient -header acc, date, invnum, ordernum, reg, lt, fuel, amt, slipno

foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods -transactiondate $aObj.date

    #Format Pastel batch
    $props = [ordered] @{
        hd    = 'Header'
        f1    = ''
        f2    = ''
        f3    = 'Y'
        acc   = $aObj.acc
        per   = $pastelper
        dte   = $aObj.date
        order = $aObj.invnum
        f4    = "N"
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
        f21   = 'Detail'
        f22   = '0'
        f23   = ''
        f24   = 'Y'
        f25   = 'Detail'
        f26   = '0'
        f27   = '1'
        f28   = $aObj.amt
        f29   = $aObj.amt
        f30   = ''
        f31   = '0'
        f32   = '0'
        f33   = '0'
        f34   = '8430000'
        f35   = 'Product' + ' : ' + $aObj.fuel
        f36   = '6'
        f37   = ''
        f38   = ''
        f39   = 'Detail'
        f40   = '0'
        f41   = '1'
        f42   = '0'
        f43   = '0'
        f44   = ''
        f45   = '0'
        f46   = '0'
        f47   = '0'
        f48   = "'"
        f49   = 'Lt' + ' : ' + $aObj.lt
        f50   = 7
        f51   = ''
        f52   = ''
        #
        f53   = 'Detail'
        f54   = '0'
        f55   = '1'
        f56   = '0'
        f57   = '0'
        f58   = ''
        f59   = '0'
        f60   = '0'
        f61   = '0'
        f62   = "'"
        f63   = 'Order' + ' : ' + $aObj.ordernum
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
        f74   = '0'
        f75   = '0'
        f76   = "'"
        f77   = 'Reg' + ' : ' + $aObj.reg
        f78   = 7
        f79   = ''
        f80   = ''
        #
        f81   = 'Detail'
        f82   = '0'
        f83   = '1'
        f84   = '0'
        f85   = '0'
        f86   = ''
        f87   = '0'
        f88   = '0'
        f89   = '0'
        f90   = "'"
        f91   = 'Slip no' + ' : ' + $aObj.slipno
        f92   = 7
        f93   = ''
        f94   = ''

    }
                 
    $objlist = New-Object -TypeName psobject -Property $props 
    $objlist | Select-Object * | Export-Csv -path $outfile -NoTypeInformation -Append
}  
#Remove header information so file can be imported into Pastel Accounting.

Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
Remove-Item -Path $outfile
#Remove-Item -Path $csvclient