#Read csv file which lists the accounts that had invoices in the reporting period.
#This file is prepared manually from Pastel Accounting report
#IF the 'CopyInv' field is True (1) then copy invoices will happen for that account
#$Datestart = '2018/12/01' must be this format
#$Dateend = '2018/12/31' must be this format
#Scheduled task 17:15 every Friday.

$playsoung = New-Object System.Media.Soundplayer
$playsoung.SoundLocation = 'C:\Users\Guy\Documents\Powershell\Sound\script start.WAV'
$playsoung.playsync()

#The Search Invoices folder contains the PDF files that have been indexed by Windows Index
$StopWatch = [System.Diagnostics.Stopwatch]::StartNew()
$searchPath = "C:\Search Invoices"

#Read the CSV file to see with customer had invoices for the reporting period
$listacc = Import-csv -path 'C:\Userdata\Circe Launches\Monthly Invoices\accountsMar2020.csv' | Where-Object { $_.CopyInv -eq 1 }
foreach ($tocopy in $listacc) {
    $customer = $($tocopy.acnum)
    $period = $($tocopy.accper)
    $mainloc = $($tocopy.destination)
    $finalcf = $($tocopy.custfolder) 
    $DateStart = $($tocopy.datestart)
    #$DateEnd = $($tocopy.dateend) 

        
    #Assemble the SQL statement for retrieving from the Windows Index
    $sql = "SELECT System.ItemName, System.ItempathDisplay, " +
    "System.DateModified FROM SYSTEMINDEX " +
    "WHERE DIRECTORY = '$searchPath' AND FREETEXT('$Customer') AND System.DateModified >= '$Datestart'"
    
    #Load the data into tables from the Windows Index (SQL database)
    $provider = "provider=search.collatordso;extended properties=’application=windows’;" 
    $connector = new-object system.data.oledb.oledbdataadapter -argument $sql, $provider 
    $dataset = new-object system.data.dataset 
    
    if ($connector.fill($dataset)) { $dataset.tables[0] | Export-Csv 'c:\Userdata\Circe Launches\Monthly Invoices\pdflist.txt' -Append }
    
    #Final location for invoices for this customer, obtained from the CSV file.
    $finalpath = "$mainloc\$finalcf\$period"    
    
    #Create the final directory for copying Offsite if it does not exist (CSV file)
    if (!(Test-Path -Path $finalpath)) {
        New-Item -Path $finalpath -ItemType Directory -Force
    }
    
    #Do the copy from the search invoices folder to Offsite
    foreach ($datarow in $dataset.Tables[0].Rows) {
        
        #Check if the files are already there, if not copy, otherewise ignore        
        if (test-path $finalpath\"$($datarow.'SYSTEM.ITEMNAME')") {
            #Write-Host 'file exists - not copying'
        }
        else {
            Copy-Item $searchpath\"$($datarow.'SYSTEM.ITEMNAME')" -Destination $finalpath -Force
        }
    }
}
$elapsed = $StopWatch.Elapsed
Write-Host "All completed. Elapsed time: $elapsed"   

$playsoung = New-Object System.Media.Soundplayer
$playsoung.SoundLocation = 'C:\Users\Guy\Documents\Powershell\Sound\script end.WAV'
$playsoung.playsync()


