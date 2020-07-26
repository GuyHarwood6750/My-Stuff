<#  Get list of Supplier invoices from spreadsheet
    Output to text file to be imported as a Pastel Supplier Invoice batch.
    
    **************  Used for MIXED VAT invoices  ******************
    #>
$inspreadsheet = 'C:\userdata\route 62\_all suppliers\suppliers July 2020.xlsm'
$outfile2 = 'C:\userdata\Route 62\_aLL suppliers\suppliers mixed vat inv july 2020_2.csv'
$outfileF = 'C:\userdata\route 62\_All Suppliers\Suppliers july mixed vat 2020_2.txt'     
$custsheet = 'july mixed vat'                          #Customer worksheet
$csvfile = 'suppliers mixed vat temp.csv'
$pathout = 'C:\userdata\route 62\_all suppliers\'
$startR = 5                                    #Start row
$endR = 11                                    #End Row
$startCol = 1                                    #Start Col (don't change)
$endCol = 15                                      #End Col (don't change)

$Outfile = $pathout + $csvfile

Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol -NoHeader -DataOnly | Where-Object -FilterScript {$_.P15 -ne 'done'} | Export-Csv -Path $Outfile -NoTypeInformation

# Format date column correctly
Get-ChildItem -Path $pathout -Name $csvfile
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.workbooks.Open($Outfile)
$xl.Sheets.Item('suppliers mixed vat temp').Activate()
  
$range = $xl.Range("D:D").Entirecolumn
$range.NumberFormat = 'dd/mm/yyyy'

$wb.save()
$xl.Workbooks.Close()
$xl.Quit()

Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
Remove-Item -Path $outfile

#Input from Supplier spreadsheet
#$csvsupplier = 'C:\userdata\route 62\_All Suppliers\Suppliers May mixed vat 2020.csv'
#Temp file      
#$outfile = 'C:\userdata\route 62\_All Suppliers\supplierinv.txt'
#File to be imported into Pastel        

#Remove last file imported to Pastel
$checkfile = Test-Path $outfileF
if ($checkfile) { Remove-Item $outfileF }                   

#Import latest csv from Supplier spreadsheet, MIXED VAT.
$data = Import-Csv -path $outfile2 -header alloc, acc, suppname, date, ref, invnum, descr, amt, gross

foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods -transactiondate $aObj.date

    #Process transactions lines
    $vatpercent = 15
    $vatpercent0 = 0
    $expaccVAT = '2000010'  
    $expaccNVAT = '2000012'

    [decimal] $total = $aObj.amt
    [decimal] $gross = $aObj.gross
    [decimal] $vat = [math]::Round($total - $gross,2)
    [decimal] $amtexVAT = [math]::Round($vat / 0.15,2)
    [decimal] $NoVAT = [math]::Round($gross - $amtexVAT,2)
    [decimal] $amtincVAT = [math]::Round($amtexVAT + $vat,2)

    #Process Supplier that are not 'default purchases'
    <#
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
    #>
    #Format Pastel batch
    $invoicedetails = [ordered] @{
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
    # Format VAT line
        f25   = 'Detail'
        f26   = $amtexVAT
        f27   = '1'
        f28   = $amtexVAT
        f29   = $amtincVAT
        f30   = ''
        f31   = $vatpercent
        f32   = '0'
        f33   = '0'
        f34   = $expaccVAT
        f35   = $aObj.descr
        f36   = '6'
        f37   = ''
        f38   = ''
    # Format non VAT line    
        f39  = 'Detail'
        f40  = $NoVAT
        f41  = '1'
        f42  = $NoVAT
        f43  = $NoVAT
        f44  = ''
        f45  = $vatpercent0
        f46  = '0'
        f47  = '0'
        f48 = $expaccNVAT
        f49 = $aObj.descr
        f50 = '6'
        f51 = ''
        f52 = ''
    }
                 
    $objlist = New-Object -TypeName psobject -Property $invoicedetails 
    $objlist | Select-Object * | Export-Csv -path $outfile -NoTypeInformation -Append
}  
#Remove header information so file can be imported into Pastel Accounting.
Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfileF
#Remove Temp file.
Remove-Item -Path $outfile