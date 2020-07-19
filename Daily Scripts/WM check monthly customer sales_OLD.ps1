<#
Read the sales file produced from Pastel to get the Customer numbers.
Then modify 'accounts.csv' accordingly.
1st step in final solution
#>
$datapath = 'c:\userdata\circe launches\monthly invoices\sales.txt'
$outfile = 'C:\userdata\circe launches\monthly invoices\outlist.txt'
#Delete outfile if it exists
$checkfile = Test-Path $outfile
if ($checkfile) {Remove-Item $outfile}
   
#Read Sales file and only select the "Customer : " line to extract the customer number.
$data1a = Get-Content $datapath | select-string -Pattern '"Customer : ' 
for ($i = 0; $i -lt $data1a.length; $i++) {
    $data2 = $data1a[$i].ToString()
    $separator = '"', 'Customer : ', '-', ' '
    $option = [System.StringSplitOptions]::RemoveEmptyEntries
    $data2.Split($separator, $option)[0] | Out-File -FilePath 'C:\userdata\circe launches\Monthly Invoices\outlist.txt' -Append -Force
}
