Function ([string]$searchString)
{
#$searchString = "erm"
$searchPath = "C:\Search Invoices"
$destinationpath = "C:\Test\pdf"
$sql = "SELECT System.ItemName, System.ItempathDisplay, " +
       "System.DateModified FROM SYSTEMINDEX " +
       "WHERE SCOPE = '$searchPath' AND FREETEXT('$searchstring')"
$provider = "provider=search.collatordso;extended properties=’application=windows’;" 
$connector = new-object system.data.oledb.oledbdataadapter -argument $sql, $provider 
$dataset = new-object system.data.dataset 
if ($connector.fill($dataset)) { $dataset.tables[0] | Export-Csv 'c:\test\pdflist.txt'}

$finalpath = New-Item -Path $destinationpath\$searchstring -ItemType Directory

foreach ($datarow in $dataset.Tables[0].Rows) {
      Copy-Item $searchpath\"$($datarow.'SYSTEM.ITEMNAME')" -Destination $finalpath
  }
}