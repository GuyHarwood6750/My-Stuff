<#      Extract from Customer spreadsheet the range for new invoices to be generated.
        Modify the $StartR (startrow) and $endR (endrow).
        This can only be done by eyeball as spreadsheet has historical data.
#>
$playsoung = New-Object System.Media.Soundplayer
$playsoung.SoundLocation = 'C:\Users\Guy\Documents\Powershell\Sound\script start.WAV'
$playsoung.playsync()

$inspreadsheet = 'C:\userdata\circe launches\InvWM\invsam08NovA.xlsx'
$csvfile = 'arrived.csv'
$csvfile2 = 'invsam08NovA.csv'

$pathout = 'C:\userdata\circe launches\InvWM\'
$custsheet = 'sheet1'                          #Customer worksheet
$startR = 1                                    #Start row
$endR = 5                                    #End Row
$startCol = 1                                    #Start Col (don't change)
$endCol = 12                                      #End Col (don't change)

$Outfile = $pathout + $csvfile
$Outfile2 = $pathout + $csvfile2

Import-Excel -Path $inspreadsheet -WorksheetName $custsheet -StartRow $startR -StartColumn $startCol -EndRow $endR -EndColumn $endCol | Export-Csv -Path $Outfile -NoTypeInformation -Encoding UTF8

# Format date column correctly
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

$playsoung = New-Object System.Media.Soundplayer
$playsoung.SoundLocation = 'C:\Users\Guy\Documents\Powershell\Sound\script end.WAV'
$playsoung.playsync()
