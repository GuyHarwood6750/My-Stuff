﻿$searchString = "erm"
$searchPath = "C:\Search Invoices"
$destinationpath = "c:\monthly invoices"
$datestart = '2018-11-01'
$sql = "SELECT System.ItemName, System.ItempathDisplay, " +
       "System.DateModified FROM SYSTEMINDEX " +
       "WHERE SCOPE = '$searchPath' AND FREETEXT('$searchstring') AND System.DateModified >= '$datestart'"
$provider = "provider=search.collatordso;extended properties=’application=windows’;" 
$connector = new-object system.data.oledb.oledbdataadapter -argument $sql, $provider 
$dataset = new-object system.data.dataset 
if ($connector.fill($dataset)) { $dataset.tables[0] | Export-Csv 'c:\test\pdflist.txt'}
        
    $finalpath = New-Item -Path $destinationpath\$searchstring -ItemType Directory -Force

    foreach ($datarow in $dataset.Tables[0].Rows) {
        Copy-Item $searchpath\"$($datarow.'SYSTEM.ITEMNAME')" -Destination $finalpath -Force
        }
    