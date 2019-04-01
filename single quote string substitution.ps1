$filelist = ((Get-ChildItem -path 'C:\Users\guy\documents\powershell').Name -join ', ')
'the folders / files are: {0}' -f $filelist