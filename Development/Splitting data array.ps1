$data = 'C:\test\invoices\sales.txt'
#$data1a = Get-Content $data | select-string -Pattern '"Customer : ' | Out-String
$data1a = Get-Content $data | select-string -Pattern '"Customer : ' 
for ($i = 0; $i -lt $data1a.length; $i++) {
    $data2 = $data1a[$i].ToString()
   write-host $data2
    $separator = '"', 'Customer : ', '-', ' '
    $option = [System.StringSplitOptions]::RemoveEmptyEntries
    $data3 = $data2.Split($separator, $option)
    $data4 = $data4 + $data3
}
