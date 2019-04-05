$a = Import-Csv -Path 'C:\test\Excelconvert\Book1a.csv'
$b = Import-Csv -Path 'C:\test\Excelconvert\Daily_Snapshot 2.csv'
foreach ($aObj in $A) {
    foreach ($BObj in $B) {
        if ($aObj.'Order No' -eq $BObj.GroupName) {
            $props = [ordered] @{
                Invoice        = $aObj.Reference
                Order          = $aObj.'Order No'
                Arrived        = $BObj.GroupName}
                 
            $objlist = New-Object -TypeName psobject -Property $props 
            $objlist | Select-Object invoice, order, arrived  | export-csv -path 'C:\test\Excelconvert\list1.csv' -NoTypeInformation -Append
            #Write-Output $objlist
            }
    }    
}