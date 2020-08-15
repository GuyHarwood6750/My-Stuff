<#      Extract cash vouchers from spreadsheet the range for new invoices to be generated.
        Modify the $StartR (startrow) and $endR (endrow).
                This can only be done by eyeball as spreadsheet has historical data.
#>
$inspreadsheet = 'C:\userdata\route 62\_All Suppliers\Suppliers august 2020.xlsm'
$outfile2 = 'C:\userdata\route 62\_All Suppliers\CSH august 2020_1.csv'
$custsheet = 'august 2020'                                #Month worksheet
$startR = 5                                             #Start row - do not change
$endR = 59                                              #End Row - change if necessary depending on number of purchases
$csvfile = 'SHEET1.csv'
$pathout = 'C:\userdata\route 62\_All Suppliers\'
$startCol = 1                                                                   #Start Col (don't change)
$endCol = 11                                                                     #End Col (don't change)
$filter = "CSH"
$outfile1 = 'C:\Userdata\route 62\_all suppliers\cashsupplier.txt'              #Temp file
$outfileF = 'C:\Userdata\route 62\_all suppliers\cashpurpastel.txt'             #File to be imported into Pastel             
$Outfile = $pathout + $csvfile

Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol -NoHeader -DataOnly| Where-Object -Filterscript { $_.P2 -eq $filter -and $_.P11 -ne 'done'} | Export-Csv -Path $Outfile -NoTypeInformation

Get-ChildItem -Path $pathout -Name $csvfile
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.workbooks.Open($Outfile)
$xl.Sheets.Item('sheet1').Activate()
$range = $xl.Range("d:d").Entirecolumn
$range.NumberFormat = 'dd/mm/yyyy'

$wb.save()
$xl.Workbooks.Close()
$xl.Quit()

Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
Remove-Item -Path $outfile

#Remove last file imported to Pastel
$checkfile = Test-Path $outfileF
if ($checkfile) { Remove-Item $outfilef }                   

#Import latest csv from Client spreadsheet
$data = Import-Csv -path $outfile2 -header Expacc, type, supplier, date, ref, ref2, desc, amt, vat    

foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods -transactiondate $aObj.date
    
    Switch ($aObj.Expacc) {
        ADV { $expacc = '3050000' }         
        AIRTIME { $expacc = '4600000' }         
        CLEANING { $expacc = '3250000' }         
        CWAGE { $expacc = '4401000' }         
        COMPA { $expacc = '6250010' }         
        COMPE { $expacc = '3300000' }         
        COURIER { $expacc = '3400000' }
        DONATION { $expacc = '3600000' }         
        ELEC { $expacc = '3650000' }
        EQUIP { $expacc = '2999000' }
        FUEL { $expacc = '4150001' }
        MVR { $expacc = '4150002' }
        PACKAGING { $expacc = '2000010' }
        POST { $expacc = '3400000' }
        NPUR { $expacc = '2000012' }
        PUR {$expacc = '2000010'}         
        PVT { $expacc = '5201001' }         
        RM { $expacc = '4350000' } 
        STATIONERY { $expacc = '4200000' }
        TEL { $expacc = '4600000' }         
        
        Default {$expacc = '9983000'}       
    }

    Switch ($aObj.vat) {
        Y { $VATind = '15' }
        N { $VATind = '0' }
        Default {$VATind = '15'}
    }
    #Format Pastel batch   
    $props1 = [ordered] @{
        Period  = $pastelper
        Date    = $aObj.date
        GL      = 'G'
        contra  = $expacc                       #Expense account to be debited (DR)
        ref     = $aObj.ref
        comment = $aObj.supplier
        amount  = $aObj.amt
        fil1    = $VATind
        fil2    = '0'
        fil3    = ' '
        fil4    = '     '
        fil5    = '8430000'                     #Cash voucher contra account number
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
    #Remove header information so file can be imported into Pastel Accounting.
    Get-Content -Path $outfile1 | Select-Object -skip 1 | Set-Content -path $outfilef
    Remove-Item -Path $outfile1