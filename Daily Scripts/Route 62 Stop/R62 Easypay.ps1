<#  Get list of Easypay Batches
    Output to text file to be imported as a Pastel Journal batch.

#>
$csvclient = 'C:\userdata\route 62\easypay\easypay736.csv'                  #Input from Client spreadsheet
$outfile1 = 'C:\userdata\route 62\easypay\easypay1.txt'                  #Temp file
$outfileF = 'C:\userdata\route 62\easypay\easypaypastel.txt'             #File to be imported into Pastel

$checkfile = Test-Path $outfileF
if ($checkfile) { Remove-Item $outfilef }                   #Remove last file imported to Pastel

#Import latest csv from Client spreadsheet
$data = Import-Csv -path $csvclient -header date, batch, approved, starcard    

foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods -transactiondate $aObj.date

    [decimal]$approved = $aObj.approved                         #Approved transactions (DR)
    [decimal]$starcard = $aObj.starcard                         #Starcard transactions (DR)
    [decimal]$nett = ($approved + $starcard) * -1               #Calculate Nett value (CR)
    
    #Format Pastel batch
    
    $props1 = [ordered] @{
        Period  = $pastelper
        Date    = $aObj.date
        GL      = 'G'
        contra  = '8430000'
        ref     = $aObj.batch
        comment = 'Easypay Sales'
        amount  = $Nett
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
        amt2    = $Nett
    }
        $props2 = [ordered] @{    
            Period = $pastelper
            Date = $aObj.date
            GL = 'G'
            contra = '9981000'
            ref = $aObj.batch
            comment = 'Easypay Sales'
            amount = $approved
            fil1 = '0'
            fil2 = '0'
            fil3 = ' '
            fil4 = '     '
            fil5 = '0000000'
            fil6 = '1'
            fil7 = '1'
            fil8 = '0'
            fil9 = '0'
            fil10 = '0'
            amt2 = $approved
        }
        $props3 = [ordered] @{
            Period = $pastelper
            Date = $aObj.date
            GL = 'G'
            contra = '9521002'
            ref = $aObj.batch
            comment = 'Star Cards'
            amount = $starcard
            fil1 = '0'
            fil2 = '0'
            fil3 = ' '
            fil4 = '     '
            fil5 = '0000000'
            fil6 = '1'
            fil7 = '1'
            fil8 = '0'
            fil9 = '0'
            fil10 = '0'
            amt2 = $starcard
        }
      
        $objlist = New-Object -TypeName psobject -Property $props1
        $objlist | Select-Object * | Export-Csv -path $outfile1 -NoTypeInformation -Append
        $objlist2 = New-Object -TypeName psobject -Property $props2
        $objlist2 | Select-Object * | Export-Csv -path $outfile1 -NoTypeInformation -Append
        $objlist3 = New-Object -TypeName psobject -Property $props3
        $objlist3 | Select-Object * | Export-Csv -path $outfile1 -NoTypeInformation -Append
    
    }  
    #Remove header information so file can be imported into Pastel Accounting.
    
    Get-Content -Path $outfile1 | Select-Object -skip 1 | Set-Content -path $outfilef
    Remove-Item -Path $outfile1