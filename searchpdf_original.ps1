$searchString = "cght"
$searchPath = "C:\Search Invoices"
$sql = "SELECT System.ItemPathDisplay, System.DateModified, " +
       "System.Size, System.FileExtension FROM SYSTEMINDEX " +
       "WHERE SCOPE = '$searchPath' AND FREETEXT('$searchstring')"
$provider = "provider=search.collatordso;extended properties=’application=windows’;" 
$connector = new-object system.data.oledb.oledbdataadapter -argument $sql, $provider 
$dataset = new-object system.data.dataset 
if ($connector.fill($dataset)) { $dataset.tables[0] }