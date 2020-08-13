#get-childitem -path '\\wserver\wmarine\management\julia private\JEFT\2019' -Recurse | Select-Object directoryname, name, Attributes | Sort-Object name -CaseSensitive | Export-Csv '\\wserver\wmarine\management\julia private\jeft\folders_2019_1.csv' 
#
get-childitem -path '\\wserver\wmarine\management\julia private' -Recurse | Select-Object FullName, name, Attributes | Where-Object -Filterscript {$_.Attributes -eq 'Directory' } | Sort-Object Fullname -CaseSensitive | Export-Csv '\\wserver\wmarine\management\julia private\jeft\_Notes\folders_je_private.csv' 
#
#get-childitem -path '\\wserver\wmarine\management\julia private\JEFT\2019' -Recurse | Sort-Object name -CaseSensitive| Export-Csv '\\wserver\wmarine\management\julia private\jeft\folders3.csv' 
#
#Where-Object -Filterscript { $_.P2 -ne $filter -and $_.P2 -ne 'CC' -and $_.P11 -ne 'done'}