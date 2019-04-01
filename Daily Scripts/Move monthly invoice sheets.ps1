    $status = Test-Connection wserver

 if ($status.statuscode -eq 0) {

    $src1 = "\\wserver\wmarine\Booking Reports\Guy"
    $dest1 = "\\wserver\wmarine\booking reports\monthly invoice reports"
    $dest2 = "\\wserver\wmarine\booking reports\Guy\OLD"
    $src2 = "\\wserver\wmarine\Booking reports\monthly invoice reports\"

        Get-ChildItem -Path $src1\*_wk.xlsx | Move-Item -Destination $dest1 -Force

        Get-ChildItem -Path $src1\*.xlsx | Move-Item -Destination $dest2 -Force

    $files = Get-ChildItem -path $src2 -filter '*.xlsx'

    Foreach ($file in $files) {
        if ($file.isreadonly -eq $false) { 
            Set-ItemProperty -path $file.FullName -Name Isreadonly -value $true  
        }
        else {
            #Get-ItemProperty -Path $src\$file | select * | Format-list
        }
    }
     }
else 
    {
        Guy-SendGmail "Move of invoice worksheets files to location failed" "PLEASE INVESTIGATE"
    } 