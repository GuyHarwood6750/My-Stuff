<#
Read the sales file produced from Pastel to get the Customer numbers.
Then modify 'accounts.csv' accordingly.
#>
$DevPath = 'C:\Users\Guy\Documents\Powershell\Development\DATA\'
$datapath = $DevPath +'sales.txt'                     #File from Pastel Accounting.
$outfile1 = $DevPath +'outlist1.txt'                  #List of Customers and their invoices for the period.
$outfile2 = $DevPath +'outlist2.txt'                  #List of Customer names only.
$pattern1 = '"Customer : '


(get-content $datapath) | Where-Object { -not $_::IsNullorWhiteSpace($_)} | Select-Object -skip 4 | Set-Content $outfile1   #Get rid of blank lines
get-Content $outfile1 | select-string -Pattern $pattern1 | Set-Content $outfile2

foreach ($lines in (get-content $outfile1)) {
    #$invoiceN = (Select-string $lines -pattern '"Doc No : ')
    write-host "a-Line $lines"
   # write-host "invoice line is $invoiceN" 
}