<#
Read the sales file produced from Pastel to get the Customer numbers.
Then modify 'accounts.csv' accordingly.
#>
$datapath = 'c:\userdata\circe launches\monthly invoices\sales.txt'                     #File from Pastel Accounting.
$outfile1 = 'C:\userdata\circe launches\monthly invoices\outlist1.txt'                  #List of Customers and their invoices for the period.
$outfile2 = 'C:\userdata\circe launches\monthly invoices\outlist2.txt'                  #List of Customer names only.
$pattern1 = '"Customer : '

(get-content $datapath) | ? { -not $_::IsNullorWhiteSpace($_)} | Set-Content $outfile1   #Get rid of blank lines
Get-Content $outfile1 | select-string -Pattern $pattern1 | Set-Content $outfile2
