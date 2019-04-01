$listacc = Import-csv -path 'C:\Monthly Invoices\M2.csv' | Where-Object {$_.CopyInv -eq 1}
foreach ($tocopy in $listacc) {
$customer = $($tocopy.acnum)
$period = $($tocopy.accper)
$mainloc = $($tocopy.destination)
$finalcf = $($tocopy.custfolder)

#Write-Host $customer, $period, $mainloc, $finalcf

#$searchString = $customer
$searchPath = "C:\Search Invoices"
$destinationpath = "c:\monthly invoices"
$datestart = '2018-11-01'
$sql = "SELECT System.ItemName, System.ItempathDisplay, " +
       "System.DateModified FROM SYSTEMINDEX " +
       "WHERE SCOPE = '$searchPath' AND FREETEXT('$customer') AND System.DateModified >= '$datestart'"
$provider = "provider=search.collatordso;extended properties=’application=windows’;" 
$connector = new-object system.data.oledb.oledbdataadapter -argument $sql, $provider 
$dataset = new-object system.data.dataset 
if ($connector.fill($dataset)) { $dataset.tables[0] | Export-Csv 'c:\test\pdflist.txt'}
        
    $finalpath = New-Item -Path $destinationpath\$finalcf\$period -ItemType Directory -Force

    foreach ($datarow in $dataset.Tables[0].Rows) {
        Copy-Item $searchpath\"$($datarow.'SYSTEM.ITEMNAME')" -Destination $finalpath -Force
        }
 }   