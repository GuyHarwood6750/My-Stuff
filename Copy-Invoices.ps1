    #Read csv file which lists the accounts that had invoices in the reporting period.
    #This file is prepared manually from Pastel Accounting report
    #IF the 'CopyInv' field is True (1) then copy invoices will happen for that account
    #$Datestart = '2018/12/01' must be this format
    #$Dateend = '2018/12/31' must be this format
    #Scheduled task 18:00 every Friday.


    $listacc = Import-csv -path 'C:\Userdata\Circe Launches\Monthly Invoices\accounts.csv' | Where-Object {$_.CopyInv -eq 1}
        foreach ($tocopy in $listacc) {
            $customer = $($tocopy.acnum)
            $period = $($tocopy.accper)
            $mainloc = $($tocopy.destination)
            $finalcf = $($tocopy.custfolder) 
            $DateStart = $($tocopy.datestart)
            $DateEnd = $($tocopy.dateend) 

    #The Search Invoices folder contains the PDF files that have been indexed by Windows Index
    $searchPath = "C:\Search Invoices"
    
    #Assemble the SQL statement for retrieving from the Windows Index
        $sql = "SELECT System.ItemName, System.ItempathDisplay, " +
        "System.DateModified FROM SYSTEMINDEX " +
        "WHERE SCOPE = '$searchPath' AND FREETEXT('$Customer') AND System.DateModified >= '$Datestart' AND System.DateModified <= '$Dateend'"
    
    #Load the data into tables from the Windows Index (SQL database)
    $provider = "provider=search.collatordso;extended properties=’application=windows’;" 
    $connector = new-object system.data.oledb.oledbdataadapter -argument $sql, $provider 
    $dataset = new-object system.data.dataset 
    
    if ($connector.fill($dataset)) {$dataset.tables[0] | Export-Csv 'c:\Userdata\Circe Launches\Monthly Invoices\pdflist.txt' -Append}
        
    #Create the final directory for copying Offsite
    $finalpath = New-Item -Path $mainloc\$finalcf\$period -ItemType Directory -Force

    #Do the copy from the search invoices folder to Offsite
    foreach ($datarow in $dataset.Tables[0].Rows) {
        Copy-Item $searchpath\"$($datarow.'SYSTEM.ITEMNAME')" -Destination $finalpath -Force
        }
    }  