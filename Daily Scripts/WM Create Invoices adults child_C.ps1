<#      Extract from Customer spreadsheet the range for new invoices to be generated.
        Modify the $endR (endrow) only.
        This can only be done by eyeball as spreadsheet has historical data.
#>
$playsoung = New-Object System.Media.Soundplayer
$playsoung.SoundLocation = 'C:\Users\Guy\Documents\Powershell\Sound\script start.WAV'
$playsoung.playsync()

$inspreadsheet = 'C:\userdata\circe launches\InvWM\inv23MarA.xlsx'
$csvfile = 'arrived.csv'                                        
$csvfile2 = 'inv23MarA.csv'
$pathout = 'C:\userdata\circe launches\InvWM\'
$custsheet = 'sheet1'                          #Customer worksheet
$startR = 1                                   #Start row (don't change)
$endR = 3                                #End Row
$startCol = 1                                    #Start Col (don't change)
$endCol = 12                                      #End Col (don't change)
#
$Outfile = $pathout + $csvfile
$Outfile2 = $pathout + $csvfile2
#
Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol | Export-Csv -Path $Outfile -NoTypeInformation -Encoding UTF8
#Format date column correctly
Get-ChildItem -Path $pathout -Name $csvfile
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$wb = $xl.workbooks.Open($Outfile)
#$xl.Sheets.Item('arrived').Activate()
$range = $xl.Range("b:b").Entirecolumn
$range.NumberFormat = 'dd/mm/yyyy'
$wb.save()
$xl.Workbooks.Close()
$xl.Quit()
#>     
Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
Remove-Item -Path $outfile
<#  
    Get list of invoices from spreadsheet
    Output to text file to be imported as a Pastel Invoice batch.
#>
$csvclient = 'c:\userdata\circe launches\invwm\inv23MarA.csv' #Input csv file
$csvrate = 'c:\userdata\circe launches\invwm\rate file\accitemrate.csv'  #Rates per customer
$outfile = 'c:\userdata\circe launches\invwm\WMinvTmp.txt'     #Temp file
$outfile2 = 'c:\userdata\circe launches\invwm\WMinv23MarA.txt'  #File to be imported into Pastel
#Remove last file imported to Pastel
if (Test-Path $outfile2) { Remove-Item $outfile2 }
#Import latest csv from Client spreadsheet
$data = Import-Csv -path $csvclient -header accnum, date, time, customername, BookingID, groupname, voucher, ptype, type, qty, rate, guide

$ratelkup = Import-csv -Path $csvrate -Header acc, code, desc, rate

$prevbookingID = 0

foreach ($aObj in $data) {
    #Return Pastel accounting period based on the transaction date.
    $pastelper = PastelPeriods -transactiondate $aObj.date
    
    # If booking id has changed then add a header record,
    # this with happen for the first header as well
    if ($aObj.BookingID -ne $prevbookingID) {
        $prevbookingID = $aObj.BookingID

        $headerProperties = [ordered] @{
            Col1  = 'Header'
            Col2  = ''
            Col3  = ''
            Col4  = 'Y'
            Col5  = $aObj.accnum
            Col6  = $pastelper
            Col7  = $aObj.date
            Col8  = $aObj.groupname
            Col9  = "Y"
            Col10 = '0'
            Col11 = ''
            Col12 = ''
            Col13 = ''
            Col14 = ''
            Col15 = ''
            Col16 = ''
            Col17 = ''
            Col18 = ''
            Col19 = ''
            Col20 = '0'
            Col21 = $aObj.date
            Col22 = ''
            Col23 = ''
            Col24 = ''
            Col25 = '1'
            Col26 = ''
            Col27 = ''
            Col28 = ''
            Col29 = 'Y'
        }
        $Line1Properties = [ordered] @{    
            Col1  = 'Detail'
            Col2  = '0'
            Col3  = '1'
            Col4  = '0'
            Col5  = '0'
            Col6  = ''
            Col7  = '0'
            Col8  = '3'
            Col9  = '0'
            Col10 = "'"
            Col11 = $aObj.groupname
            Col12 = 7
            Col13 = ''
            Col14 = ''
            Col15 = ''
            Col16 = ''
            Col17 = ''
            Col18 = ''
            Col19 = ''
            Col20 = ''
            Col21 = ''
            Col22 = ''
            Col23 = ''
            Col24 = ''
            Col25 = ''
            Col26 = ''
            Col27 = ''
            Col28 = ''
            Col29 = '' 
        }
        #
        $Line2Properties = [ordered] @{
            col1  = 'Detail'
            Col2  = '0'
            Col3  = '1'
            Col4  = '0'
            Col5  = '0'
            Col6  = ''
            Col7  = '0'
            Col8  = '3'
            Col9  = '0'
            col10 = "'"
            Col11 = $aObj.voucher
            col12 = 7
            col13 = ''
            col14 = ''
            col15 = ''
            col16 = ''
            col17 = ''
            col18 = ''
            col19 = ''
            col20 = ''
            col21 = ''
            col22 = ''
            col23 = ''
            col24 = ''
            col25 = ''
            col26 = ''
            col27 = ''
            col28 = ''
            col29 = ''
        }
        $Line3Properties = [ordered] @{
            col1  = 'Detail'
            Col2  = '0'
            Col3  = '1'
            Col4  = '0'
            Col5  = '0'
            Col6  = ''
            Col7  = '0'
            Col8  = '3'
            Col9  = '0'
            col10 = "'"
            Col11 = 'Date: ' + ' ' + $aObj.date
            col12 = 7
            col13 = ''
            col14 = ''
            col15 = ''
            col16 = ''
            col17 = ''
            col18 = ''
            col19 = ''
            col20 = ''
            col21 = ''
            col22 = ''
            col23 = ''
            col24 = ''
            col25 = ''
            col26 = ''
            col27 = ''
            col28 = ''
            col29 = ''
        }
        $Line4Properties = [ordered] @{
            col1  = 'Detail'
            Col2  = '0'
            Col3  = '1'
            Col4  = '0'
            Col5  = '0'
            Col6  = ''
            Col7  = '0'
            Col8  = '3'
            Col9  = '0'
            col10 = "'"
            Col11 = 'Time: ' + ' ' + $aObj.time
            col12 = 7
            col13 = ''
            col14 = ''
            col15 = ''
            col16 = ''
            col17 = ''
            col18 = ''
            col19 = ''
            col20 = ''
            col21 = ''
            col22 = ''
            col23 = ''
            col24 = ''
            col25 = ''
            col26 = ''
            col27 = ''
            col28 = ''
            col29 = ''
        }
        $Line5Properties = [ordered] @{
            col1  = 'Detail'
            Col2  = '0'
            Col3  = '1'
            Col4  = '0'
            Col5  = '0'
            Col6  = ''
            Col7  = '0'
            Col8  = '3'
            Col9  = '0'
            col10 = "'"
            Col11 = 'Guide: ' + ' ' + $aObj.guide
            col12 = 7
            col13 = ''
            col14 = ''
            col15 = ''
            col16 = ''
            col17 = ''
            col18 = ''
            col19 = ''
            col20 = ''
            col21 = ''
            col22 = ''
            col23 = ''
            col24 = ''
            col25 = ''
            col26 = ''
            col27 = ''
            col28 = ''
            col29 = ''
        }
        $Line6Properties = [ordered] @{
            col1  = 'Detail'
            Col2  = '0'
            Col3  = '1'
            Col4  = '0'
            Col5  = '0'
            Col6  = ''
            Col7  = '0'
            Col8  = '3'
            Col9  = '0'
            col10 = "'"
            Col11 = 'Our ref: ' + ' ' + $aObj.BookingID
            col12 = 7
            col13 = ''
            col14 = ''
            col15 = ''
            col16 = ''
            col17 = ''
            col18 = ''
            col19 = ''
            col20 = ''
            col21 = ''
            col22 = ''
            col23 = ''
            col24 = ''
            col25 = ''
            col26 = ''
            col27 = ''
            col28 = ''
            col29 = ''
        }
        $Line7Properties = [ordered] @{
            col1  = 'Detail'
            Col2  = '0'
            Col3  = '1'
            Col4  = '0'
            Col5  = '0'
            Col6  = ''
            Col7  = '0'
            Col8  = '3'
            Col9  = '0'
            col10 = "'"
            Col11 = ' '
            col12 = 7
            col13 = ''
            col14 = ''
            col15 = ''
            col16 = ''
            col17 = ''
            col18 = ''
            col19 = ''
            col20 = ''
            col21 = ''
            col22 = ''
            col23 = ''
            col24 = ''
            col25 = ''
            col26 = ''
            col27 = ''
            col28 = ''
            col29 = ''
        }
        # Append the header and invoice lines to the CSV file
        $objHeader = New-Object -TypeName psobject -Property $headerProperties 
        $objHeader | Select-Object * | Export-Csv -path $outfile -Append -NoTypeInformation

        $objGroup = New-Object -TypeName psobject -Property $Line1Properties 
        $objGroup | Select-Object * | Export-Csv -path $outfile -Append -NoTypeInformation

        $objGroup = New-Object -TypeName psobject -Property $Line2Properties 
        $objGroup | Select-Object * | Export-Csv -path $outfile -Append -NoTypeInformation
        
        $objGroup = New-Object -TypeName psobject -Property $Line3Properties 
        $objGroup | Select-Object * | Export-Csv -path $outfile -Append -NoTypeInformation
        
        $objGroup = New-Object -TypeName psobject -Property $Line4Properties 
        $objGroup | Select-Object * | Export-Csv -path $outfile -Append -NoTypeInformation
        
        $objGroup = New-Object -TypeName psobject -Property $Line5Properties 
        $objGroup | Select-Object * | Export-Csv -path $outfile -Append -NoTypeInformation
        
        $objGroup = New-Object -TypeName psobject -Property $Line6Properties 
        $objGroup | Select-Object * | Export-Csv -path $outfile -Append -NoTypeInformation
        
        $objGroup = New-Object -TypeName psobject -Property $Line7Properties 
        $objGroup | Select-Object * | Export-Csv -path $outfile -Append -NoTypeInformation
    }

    foreach ($bObj in $ratelkup) {
        if ($aObj.accnum -eq $bObj.acc -and ($aobj.type -eq $bObj.desc)) {
            $ratecode = $bObj.code
            #$description = $bObj.desc
            [decimal]$amount = $aObj.rate                   #$bObj.rate (rate file csv) otherwise rate from spreadsheet
            [decimal]$vat = $amount * 15 / 115
            [decimal]$amtexvat = $aObj.rate - $vat          #$bObj.rate (rate file csv) otherwise rate from spreadsheet
            $vatexamt = [math]::Round($amtexvat, 2)
        }
        else {
        }
    }
    #Add the current row to the objects
    $detailProperties = [ordered] @{
        Col1  = 'Detail'
        Col2  = '0'
        Col3  = $aObj.qty
        Col4  = $vatexamt
        Col5  = $amount
        Col6  = ''
        Col7  = '15'
        Col8  = '3'
        Col9  = '0'
        Col10 = $ratecode                    #rate code
        Col11 = $aObj.type
        Col12 = '4'
        Col13 = ''
        Col14 = ''
        Col15 = 001
        Col16 = ''
        Col17 = ''
        Col18 = ''
        Col19 = ''
        Col20 = ''
        Col21 = ''
        Col22 = ''
        Col23 = ''
        Col24 = ''
        Col25 = ''
        Col26 = ''
        Col27 = ''
        Col28 = ''
        Col29 = ''
    } 

    $objDetails = New-Object -TypeName psobject -Property $detailProperties 
    $objDetails | Select-Object * | Export-Csv -path $outfile -Append -NoTypeInformation
}  
#Remove header information so file can be imported into Pastel Accounting.

Get-Content -Path $outfile | Select-Object -skip 1 | Set-Content -path $outfile2
Remove-Item -Path $outfile
#Remove-Item -Path $csvclient
$playsoung = New-Object System.Media.Soundplayer
$playsoung.SoundLocation = 'C:\Users\Guy\Documents\Powershell\Sound\script end.WAV'
$playsoung.playsync()
