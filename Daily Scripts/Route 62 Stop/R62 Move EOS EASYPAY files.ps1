$src1 = 'C:\Users\Guy\Dropbox\R62\Accounts\EOS'
$dest1 = 'C:\Userdata\Route 62\EOS Not Processed'
$src2 = 'C:\Users\Guy\Dropbox\R62\Accounts\EASYPAY'
$dest2 = 'C:\Userdata\Route 62\Easypay'
$dest3 = 'C:\Userdata\Route 62\EOS Not Processed\Reports'

Get-ChildItem -Path $src1\* -include *.txt | Move-Item -Destination $dest1 -Force
Get-ChildItem -Path $src1\* -include *.pdf | Move-Item -Destination $dest3 -Force
Get-ChildItem -Path $src2\* -include *.pdf | Move-Item -Destination $dest2 -Force