<#      Extract cash vouchers from spreadsheet the range for new invoices to be generated.
        Modify the $StartR (startrow) and $endR (endrow).
                This can only be done by eyeball as spreadsheet has historical data.
#>
$inspreadsheet = 'C:\userdata\circe launches\_All Suppliers\Supplier invoices cash vouchers 2021.xlsx'
$outfile2 = 'C:\userdata\circe launches\_All Suppliers\CSH July 2021_2.csv'
$custsheet = 'July 2020'                                #Month worksheet
$startR = 2                                             #Start row - do not change
$endR = 46                                              #End Row - change if necessary depending on number of purchases
$csvfile = 'SHEET1.csv'
$pathout = 'C:\userdata\circe launches\_All Suppliers\'
$startCol = 1                                                                   #Start Col (don't change)
$endCol = 10                                                                     #End Col (don't change)
$filter = "CSH"
$outfile1 = 'C:\Userdata\circe launches\_all suppliers\cashsupplier.txt'              #Temp file
$outfileF = 'C:\Userdata\circe launches\_all suppliers\cashpurpastel.txt'             #File to be imported into Pastel             
$Outfile = $pathout + $csvfile

Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol -NoHeader -DataOnly| Where-Object -Filterscript { $_.P1 -eq $filter -and $_.p9 -eq 'cash'-and $_.P10 -ne 'done'} | Export-Csv -Path $Outfile -NoTypeInformation

Get-ChildItem -Path $pathout -Name $csvfile
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.workbooks.Open($Outfile)
$xl.Sheets.Item('sheet1').Activate()
$range = $xl.Range("c:c").Entirecolumn
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
$data = Import-Csv -path $outfile2 -header acc, Expacc, date, ref, invnum, desc, amt, vat     

foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods -transactiondate $aObj.date
    
    Switch ($aObj.Expacc) {
        CLN { $expacc = '3210000' }            #Cleaning
        FUEL { $expacc = '4150000' }         #Motor vehicles
        GIFT { $expacc = '3551000' }            #Trade Gifts
        MED { $expacc = '4500000' }            #Medical expenses, Staff welfare
        MVE { $expacc = '4150000' }         #Motor vehicles
        PC { $expacc = '4550000' }            #Protective clothing
        RENT { $expacc = '4300000' }            #Rent
        RM { $expacc = '4350000' }            #Repairs and Maintenance
        REF { $expacc = '4500000' }            #Staff refreshments
        SS { $expacc = '3750000' }            #Ship stores & provisions
        STATIONARY { $expacc = '4200000' }    #Stationery
        TEL { $expacc = '4600000' }            #Telephone
        TETA { $expacc = '4451000' }            #TETA Training
        Default { $expacc = '9992000' }       #Unallocated Expense account      
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
        comment = $aObj.desc
        amount  = $aObj.amt
        fil1    = $VATind
        fil2    = '0'
        fil3    = ' '
        fil4    = '     '
        fil5    = '8410000'                     #Cash voucher contra account number
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