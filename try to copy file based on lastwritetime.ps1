Get-ChildItem "D:\Pastel Backups - Hout Bay\2019" | Where-Object {$_.LastWriteTime -gt "09/11/2018 05:00 PM"} | copy-item -Destination "c:\test" -WhatIf
$srcfile2 = Get-ItemProperty -Path "D:\Pastel Backups - Hout Bay\2019\circe19a(20181110*" -Name Lastwritetime
$srcfile = Get-ChildItem "D:\Pastel Backups - Hout Bay\2019\circe19a(201811*" | Where-Object {$_.LastWriteTime -ge "10/11/2018"}