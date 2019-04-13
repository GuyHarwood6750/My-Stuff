<#  Get list of Nedbank batches from Client.
    Output to text file to be imported as a Pastel Journal batch.

#>
$pastelperiod = 2                                           #MODIFY THIS BASED ON PASTEL PERIOD
$csvclient = 'C:\userdata\route 62\Nedbank\apr 19.csv'      #Input from Client spreadsheet
$outfile = 'C:\userdata\Route 62\nedbank\nedpas.txt'
$outfile2 = 'C:\userdata\Route 62\nedbank\nedpas2.txt'      #File to be imported into Pastel

$checkfile = Test-Path $outfile2
if ($checkfile) { Remove-Item $outfile2 }


$data = Import-Csv -path $csvclient -header date, batch, amt
foreach ($aObj in $data) {
    $props = [ordered] @{
        Period  = $pastelperiod
        Date    = $aObj.date
        GL      = 'G'
        contra  = '9984000'
        ref     = $aObj.batch
        comment = 'Nedbank CC'
        amount  = $aObj.amt
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
        amt2    = $aObj.amt
    }
                 
    $objlist = New-Object -TypeName psobject -Property $props 
    $objlist | Select-Object * | Export-Csv -path $outfile -NoTypeInformation -Append
}  
#Remove header information so file can be imported into Pastel Accounting.
#
Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
Remove-Item -Path $outfile