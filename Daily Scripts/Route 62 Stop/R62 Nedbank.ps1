<#  Get list of Nedbank batches from Client.
    Output to text file to be imported as a Pastel Journal batch.

#>
$csvclient = 'C:\userdata\route 62\Nedbank\nedbank66032.csv'          #Input from Client spreadsheet
$outfile = 'C:\userdata\Route 62\nedbank\nedpas.txt'            #Temp file
$outfile2 = 'C:\userdata\Route 62\nedbank\nedbankbatch.txt'     #File to be imported into Pastel

$checkfile = Test-Path $outfile2
if ($checkfile) { Remove-Item $outfile2 }                   #Remove last file imported to Pastel


$data = Import-Csv -path $csvclient -header date, batch, amt    #Import latest csv from Client spreadsheet

#Format Pastel batch
foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods -transactiondate $aObj.date
 
    [decimal]$DRAmt = $aObj.amt                                   #Debit value (DR)                                 
    [decimal]$CRAmt = ($DRamt) * -1                               #Credit value (CR)                              

    $props1 = [ordered] @{
        Period  = $pastelper
        Date    = $aObj.date
        GL      = 'G'
        contra  = '9984000'
        ref     = $aObj.batch
        comment = 'Nedbank CC'
        amount  = $DRAmt
        fil1    = '0'
        fil2    = '0'
        fil3    = ' '
        fil4    = '     '
        fil5    = '0000000'
        fil6    = '1'
        fil7    = '1'
        fil8    = '0'
        fil9    = '0'
        fil10   = '0'
        amt2    = $DRAmt
    }
    $props2 = [ordered] @{
        Period  = $pastelper
        Date    = $aObj.date
        GL      = 'G'
        contra  = '8430000'
        ref     = $aObj.batch
        comment = 'Nedbank CC'
        amount  = $CRAmt
        fil1    = '0'
        fil2    = '0'
        fil3    = ' '
        fil4    = '     '
        fil5    = '0000000'
        fil6    = '1'
        fil7    = '1'
        fil8    = '0'
        fil9    = '0'
        fil10   = '0'
        amt2    = $CRAmt
    }    
    $objlist1 = New-Object -TypeName psobject -Property $props1 
    $objlist1 | Select-Object * | Export-Csv -path $outfile -NoTypeInformation -Append
    $objlist2 = New-Object -TypeName psobject -Property $props2 
    $objlist2 | Select-Object * | Export-Csv -path $outfile -NoTypeInformation -Append

}  
#Remove header information so file can be imported into Pastel Accounting.
#
Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
Remove-Item -Path $outfile