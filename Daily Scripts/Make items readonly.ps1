$src = "\\wserver\wmarine\Booking reports\monthly invoice reports\"
#$src = 'C:\test\itemdetails\'

$files = Get-ChildItem -path $src -filter '*.xlsx'

 Foreach ($file in $files) {
    if ($file.isreadonly -eq $false) { 
        Set-ItemProperty -path $file.FullName -Name Isreadonly -value $true  
    }
    else {
        #Get-ItemProperty -Path $src\$file | select * | Format-list
    }
 }